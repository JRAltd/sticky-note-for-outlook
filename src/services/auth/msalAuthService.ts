import {
  AccountInfo,
  AuthenticationResult,
  PublicClientApplication,
} from "@azure/msal-browser";
import {
  DEFAULT_AUTH_CONFIG,
  LOCAL_STORAGE_AUTH_CONFIG_KEY,
} from "../../config/constants";
import { AzureAuthConfig } from "../../types/note";

const GRAPH_SCOPES = ["User.Read", "Mail.ReadWrite"];

class MsalAuthService {
  private msalInstance: PublicClientApplication | null = null;
  private initializedClientId: string = "";
  private activeAccount: AccountInfo | null = null;

  public getConfig(): AzureAuthConfig {
    try {
      const raw = localStorage.getItem(LOCAL_STORAGE_AUTH_CONFIG_KEY);
      if (raw) {
        const parsed = JSON.parse(raw) as Partial<AzureAuthConfig>;
        return {
          ...DEFAULT_AUTH_CONFIG,
          ...parsed,
        };
      }
    } catch {
      // Fallback to default config
    }
    return { ...DEFAULT_AUTH_CONFIG };
  }

  public saveConfig(newConfig: Partial<AzureAuthConfig>): AzureAuthConfig {
    const merged: AzureAuthConfig = {
      ...this.getConfig(),
      ...newConfig,
    };
    localStorage.setItem(LOCAL_STORAGE_AUTH_CONFIG_KEY, JSON.stringify(merged));
    // Reset MSAL instance if clientId changed
    if (merged.clientId !== this.initializedClientId) {
      this.msalInstance = null;
      this.activeAccount = null;
    }
    return merged;
  }

  public isGraphConfigured(): boolean {
    const cfg = this.getConfig();
    return Boolean(
      cfg.clientId &&
        cfg.clientId.trim().length > 10 &&
        cfg.clientId !== "00000000-0000-0000-0000-000000000000"
    );
  }

  private async ensureMsalInstance(): Promise<PublicClientApplication | null> {
    const cfg = this.getConfig();
    if (!this.isGraphConfigured()) {
      return null;
    }

    if (this.msalInstance && this.initializedClientId === cfg.clientId) {
      return this.msalInstance;
    }

    const authority = `https://login.microsoftonline.com/${cfg.tenantId || "common"}`;
    const instance = new PublicClientApplication({
      auth: {
        clientId: cfg.clientId.trim(),
        authority,
        redirectUri: cfg.redirectUri || window.location.origin + "/taskpane.html",
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
      },
    });

    await instance.initialize();
    const accounts = instance.getAllAccounts();
    if (accounts.length > 0) {
      this.activeAccount = accounts[0];
    }

    this.msalInstance = instance;
    this.initializedClientId = cfg.clientId;
    return instance;
  }

  /**
   * Attempts silent token acquisition via OfficeRuntime SSO first, then MSAL cache,
   * and optionally falls back to interactive popup when interactive=true.
   */
  public async getGraphAccessToken(interactive = false): Promise<string | null> {
    // 1. Try OfficeRuntime.auth SSO if running in Outlook with SSO configured
    if (
      typeof OfficeRuntime !== "undefined" &&
      OfficeRuntime.auth &&
      typeof OfficeRuntime.auth.getAccessToken === "function"
    ) {
      try {
        const ssoToken = await OfficeRuntime.auth.getAccessToken({
          allowSignInPrompt: interactive,
          allowConsentPrompt: interactive,
          forMSGraphAccess: true,
        });
        if (ssoToken) {
          return ssoToken;
        }
      } catch {
        // Office SSO not configured in manifest or not supported; fall through to MSAL
      }
    }

    // 2. Try MSAL PublicClientApplication
    const msal = await this.ensureMsalInstance();
    if (!msal) {
      return null;
    }

    const accounts = msal.getAllAccounts();
    const account = this.activeAccount || accounts[0];

    if (account) {
      try {
        const silentResult: AuthenticationResult = await msal.acquireTokenSilent({
          scopes: GRAPH_SCOPES,
          account,
        });
        this.activeAccount = silentResult.account;
        return silentResult.accessToken;
      } catch {
        // Silent token expired or requires consent; fall through if interactive
      }
    }

    if (!interactive) {
      return null;
    }

    const popupResult = await msal.loginPopup({
      scopes: GRAPH_SCOPES,
      prompt: "select_account",
    });
    this.activeAccount = popupResult.account;
    return popupResult.accessToken;
  }

  public async signOut(): Promise<void> {
    const msal = await this.ensureMsalInstance();
    if (msal && this.activeAccount) {
      await msal.logoutPopup({ account: this.activeAccount });
    }
    this.activeAccount = null;
  }

  public getConnectedAccount(): AccountInfo | null {
    return this.activeAccount;
  }
}

export const msalAuthService = new MsalAuthService();

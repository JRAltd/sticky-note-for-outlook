import { useEffect, useState } from "react";
import { DEMO_SAMPLE_EMAILS } from "../config/constants";
import { officeContextService } from "../services/office/officeContextService";
import { EmailContextInfo } from "../types/note";

export function useEmailContext() {
  const [emailContext, setEmailContext] = useState<EmailContextInfo>(
    DEMO_SAMPLE_EMAILS[0]
  );
  const [isInitializing, setIsInitializing] = useState<boolean>(true);

  useEffect(() => {
    let mounted = true;

    officeContextService.initialize().then((ctx) => {
      if (mounted) {
        setEmailContext(ctx);
        setIsInitializing(false);
      }
    });

    const unsubscribe = officeContextService.onItemChanged((updatedCtx) => {
      if (mounted) {
        setEmailContext(updatedCtx);
      }
    });

    return () => {
      mounted = false;
      unsubscribe();
    };
  }, []);

  const switchDemoEmail = (index: number) => {
    const updated = officeContextService.selectDemoEmail(index);
    setEmailContext(updated);
  };

  return {
    emailContext,
    isInitializing,
    isInOutlook: officeContextService.isInOutlook(),
    demoEmails: officeContextService.getDemoEmails(),
    switchDemoEmail,
  };
}

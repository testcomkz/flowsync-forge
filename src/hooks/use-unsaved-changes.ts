import { useEffect } from "react";
import { useBeforeUnload, useBlocker } from "react-router-dom";

/**
 * Shows a confirmation dialog when the user attempts to close the tab or
 * navigate away with pending form changes. The hook integrates with the
 * browser's beforeunload event and the client-side router navigation.
 */
export function useUnsavedChangesWarning(when: boolean, message = "You have unsaved changes. They will be lost if you leave this page.") {
  useBeforeUnload(
    event => {
      if (!when) {
        return;
      }

      event.preventDefault();
      event.returnValue = message;
      return message;
    },
    { capture: true }
  );

  const blocker = useBlocker(when);

  useEffect(() => {
    if (blocker.state !== "blocked") {
      return;
    }

    const shouldLeave = window.confirm(message);
    if (shouldLeave) {
      blocker.proceed();
    } else {
      blocker.reset();
    }
  }, [blocker, message]);
}


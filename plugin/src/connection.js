(function (scope) {
  scope.figma2pptxCreateAutoConnector = function ({probe, onConnecting, onReconnecting, onConnected, onUnauthorized, schedule = setTimeout, cancel = clearTimeout}) {
    const delays = [250, 500, 1000, 2000, 5000, 10000];
    let attempt = 0;
    let secret = '';
    let timer;
    let stopped = true;
    let inFlight = false;

    const clearTimer = () => {
      if (timer !== undefined) cancel(timer);
      timer = undefined;
    };

    const queueRetry = () => {
      if (stopped || !secret || timer !== undefined) return;
      onReconnecting();
      const delay = delays[Math.min(attempt, delays.length - 1)];
      attempt += 1;
      timer = schedule(() => {
        timer = undefined;
        void run();
      }, delay);
    };

    const run = async () => {
      if (stopped || !secret || inFlight) return;
      inFlight = true;
      try {
        const result = await probe(secret);
        if (result.unauthorized) {
          stopped = true;
          secret = '';
          clearTimer();
          onUnauthorized(result);
          return;
        }
        if (!result.ok) throw Error(result.error || 'Bridge health check failed');
        attempt = 0;
        clearTimer();
        onConnected(result);
      } catch {
        queueRetry();
      } finally {
        inFlight = false;
      }
    };

    return {
      start(value) {
        secret = String(value || '').trim();
        stopped = !secret;
        attempt = 0;
        clearTimer();
        if (!stopped) {
          onConnecting();
          return run();
        }
      },
      retry() {
        if (stopped || !secret) return;
        clearTimer();
        queueRetry();
      },
      stop() {
        stopped = true;
        secret = '';
        attempt = 0;
        clearTimer();
      },
    };
  };
})(globalThis);

// These helpers are proposed helpers for @microsoft/mgt-react

import { useEffect, useState } from 'react';
import { Providers, ProviderState } from '@microsoft/mgt';

export function useIsSignedIn() : [boolean] {
  const [isSignedIn, setIsSignedIn] = useState(false);

  useEffect(() => {
    const updateState = () => {
      let provider = Providers.globalProvider;
      setIsSignedIn(provider && provider.state === ProviderState.SignedIn);
    };

    Providers.onProviderUpdated(updateState);
    updateState();

    return () => {
      Providers.removeProviderUpdatedListener(updateState);
    }
  }, []);

  return [isSignedIn];
}

export interface GetOptions {
  version?: string,
  pollingRate?: number // TODO - poll the api at this rate - polls delta api when specified as resource
  maxPages?: number // TODO - follow pages up to the max number
  scopes?: string[] // TODO - prereq scopes to make sure are requested before making call
  
}

export function useGet<T = any>(resource: string, deps?: unknown[], options?: GetOptions) : [T | undefined, boolean, any] {
  const [response, setResponse] = useState<T>();
  const [error, setError] = useState();
  const [loading, setLoading] = useState(true);
  const [isSignedIn] = useIsSignedIn();

  useEffect(() => {
    if (isSignedIn && (!deps || deps.every(d => !!d))) {
      (async () => {
        try {
          let version = options ? (options.version ?? 'v1.0') : 'v1.0';
          setResponse(
            await Providers.globalProvider.graph.client.api(resource).version(version).get()
          );
        } catch (e) {
          setError(e);
        }
        setLoading(false);
      })();
    };

    return () => {
        // stop timers or unsubscribe from events here
    }
  }, [isSignedIn, resource]);

  // TODO, add function to return array to call when the dev wants to fetch more pages manually 
  return [response, loading, error];
}
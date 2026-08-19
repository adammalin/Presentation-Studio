/** Share one in-progress asynchronous operation for an exact stable key. */
export function singleFlight<T>(flights: Map<string, Promise<T>>, key: string, operation: () => Promise<T>): Promise<T> {
  const existing = flights.get(key);
  if (existing) return existing;
  const running = operation().finally(() => {
    if (flights.get(key) === running) flights.delete(key);
  });
  flights.set(key, running);
  return running;
}

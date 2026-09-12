/** Placeholder for modules that are scaffolded but not yet built. Every use of this must be gone before 1.0. */
export function notImplemented(module: string): Error {
  return new Error(`@jetstreamapp/simple-excel: ${module} is not implemented yet`);
}

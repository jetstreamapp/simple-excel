/**
 * Phase gates for the integration suites. Each flips to true when the corresponding facade lands, so `npm test`
 * stays green while the library is scaffolded. Both must be true (and this file deleted) before 1.0.
 */
export const WRITER_READY: boolean = false;
export const READER_READY: boolean = false;

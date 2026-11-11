import { describe, expect, it, vi } from 'vitest';
import { classifyError, isXlsxError, XlsxError } from '../errors';

describe('XlsxError', () => {
  it('keeps the (code, message, detail) call shape, with no cause unless one is given', () => {
    const error = new XlsxError('LIMIT_EXCEEDED', 'too many rows', { maxRows: 10 });
    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('XlsxError');
    expect(error.code).toBe('LIMIT_EXCEEDED');
    expect(error.message).toBe('too many rows');
    expect(error.detail).toEqual({ maxRows: 10 });
    expect('cause' in error).toBe(false);
  });

  it('puts the original error in the standard cause, not in detail', () => {
    const original = new Error('EPIPE');
    const error = new XlsxError('IO_ERROR', 'the stream failed', { path: '/tmp/x.xlsx' }, { cause: original });
    expect(error.cause).toBe(original);
    expect(error.detail).toEqual({ path: '/tmp/x.xlsx' });
  });

  it('is an instance of XlsxError from every other copy of the library', async () => {
    vi.resetModules();
    const otherCopy = (await import('../errors')) as typeof import('../errors');
    expect(otherCopy.XlsxError).not.toBe(XlsxError);

    const fromOther = new otherCopy.XlsxError('ABORTED', 'cancelled');
    expect(fromOther instanceof XlsxError).toBe(true);
    expect(isXlsxError(fromOther)).toBe(true);
    const fromThis = new XlsxError('IO_ERROR', 'failed');
    expect(fromThis instanceof otherCopy.XlsxError).toBe(true);
    expect(otherCopy.isXlsxError(fromThis)).toBe(true);
  });

  it('is not claimed by an ordinary error, a lookalike object or a primitive', () => {
    const lookalike = { name: 'XlsxError', code: 'ABORTED', message: 'not really' };
    for (const value of [new Error('plain'), lookalike, null, undefined, 'XlsxError', 42]) {
      expect(value instanceof XlsxError).toBe(false);
      expect(isXlsxError(value)).toBe(false);
    }
  });

  it('keeps the ordinary prototype check for a subclass', () => {
    class SpecialError extends XlsxError {}
    expect(new SpecialError('ABORTED', 'x') instanceof SpecialError).toBe(true);
    expect(new XlsxError('ABORTED', 'x') instanceof SpecialError).toBe(false);
    expect(new SpecialError('ABORTED', 'x') instanceof XlsxError).toBe(true);
  });
});

describe('classifyError', () => {
  it('passes an XlsxError through and wraps anything else with it as the cause', () => {
    const classified = new XlsxError('LIMIT_EXCEEDED', 'too big');
    expect(classifyError(classified, 'IO_ERROR', 'failed')).toBe(classified);
    const original = new TypeError('boom');
    const wrapped = classifyError(original, 'IO_ERROR', 'failed');
    expect(wrapped).toMatchObject({ code: 'IO_ERROR', message: 'failed', cause: original });
    expect(wrapped.detail).toBeUndefined();
  });
});

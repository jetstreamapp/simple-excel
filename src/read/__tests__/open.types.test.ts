/**
 * Type-level contract of `openWorkbook` and `Sheet.rows`, checked by `npm run typecheck` (tsconfig.typecheck.json
 * includes the tests). Every call sits in a closure that is never invoked: `expectTypeOf` only looks at its type.
 */
import { describe, expectTypeOf, it } from 'vitest';
import type { CellError, CellValue, OpenOptions, ReadValue, Workbook } from '../../types';
import { openWorkbook } from '../open';

declare const bytes: Uint8Array;
declare const workbook: Workbook;

describe('openWorkbook overloads', () => {
  it('types rows as CellValue by default and for a literal errors: string or null', () => {
    expectTypeOf(() => openWorkbook(bytes)).returns.resolves.toEqualTypeOf<Workbook<CellValue>>();
    expectTypeOf(() => openWorkbook(bytes, {})).returns.resolves.toEqualTypeOf<Workbook<CellValue>>();
    expectTypeOf(() => openWorkbook(bytes, { errors: 'string' })).returns.resolves.toEqualTypeOf<Workbook<CellValue>>();
    expectTypeOf(() => openWorkbook(bytes, { errors: 'null', dates: 'utc' })).returns.resolves.toEqualTypeOf<Workbook<CellValue>>();
  });

  it("includes CellError for a literal errors: 'object'", () => {
    expectTypeOf(() => openWorkbook(bytes, { errors: 'object' })).returns.resolves.toEqualTypeOf<Workbook<CellValue | CellError>>();
  });

  it('includes CellError when errors is not a literal', () => {
    const options: OpenOptions = { errors: 'object' };
    expectTypeOf(() => openWorkbook(bytes, options)).returns.resolves.toEqualTypeOf<Workbook<ReadValue>>();
    const errors = options.errors;
    expectTypeOf(() => openWorkbook(bytes, { errors })).returns.resolves.toEqualTypeOf<Workbook<ReadValue>>();
  });
});

describe('Sheet.rows overloads', () => {
  it('yields arrays by default and records for mode: object', () => {
    expectTypeOf(() => workbook.sheet(0).rows()).returns.toEqualTypeOf<AsyncIterable<CellValue[]>>();
    expectTypeOf(() => workbook.sheet(0).rows({ mode: 'array' })).returns.toEqualTypeOf<AsyncIterable<CellValue[]>>();
    expectTypeOf(() => workbook.sheet(0).rows({ mode: 'object', headerRow: 2 })).returns.toEqualTypeOf<
      AsyncIterable<Record<string, CellValue>>
    >();
  });

  it('accepts a mode only known at run time and yields either shape', () => {
    const mode = Math.random() > 0.5 ? 'array' : 'object';
    expectTypeOf(() => workbook.sheet(0).rows({ mode })).returns.toEqualTypeOf<AsyncIterable<CellValue[] | Record<string, CellValue>>>();
  });

  it('still refuses an unknown mode', () => {
    expectTypeOf(() => {
      // @ts-expect-error -- 'csv' is not a row mode
      return workbook.sheet(0).rows({ mode: 'csv' });
    }).toBeFunction();
  });
});

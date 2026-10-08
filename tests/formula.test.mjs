import test from 'node:test';
import assert from 'node:assert/strict';
import { formulaEvaluator, FormulaError, parseFormula, rebaseFormula } from '../src/sheets/formula.js';

function evaluator(entries = {}, options = {}) {
  return formulaEvaluator({ names: ['Orders', 'Summary'], readCell: async (sheet, address) => entries[sheet + ':' + address.row + ':' + address.column], ...options });
}

test('Calc operator precedence, left-associative power and decimal arithmetic are explicit', async () => {
  const engine = evaluator();
  for (const [formula, expected] of [['2+3*4', 14], ['-2^2', 4], ['2^3^2', 64], ['(2+3)*4', 20], ['50%*4', 2], ['ROUND(1.005,2)', 1.01], ['ROUND(-2.675,2)', -2.68], ['MOD(-7,3)', 2], ['"Word"&" "&42', 'Word 42'], ['"A"="a"', true], ['TRUE()', true]]) assert.equal(await engine.evaluate(formula), expected, formula);
  engine.dispose();
});

test('conditions evaluate lazily and propagate errors without executing document code', async () => {
  const engine = evaluator();
  assert.equal(await engine.evaluate('IF(TRUE,42,1/0)'), 42);
  assert.equal(await engine.evaluate('IF(FALSE,1/0)'), false);
  assert.equal(await engine.evaluate('IFERROR(1/0,"empty")'), 'empty');
  assert.equal(await engine.evaluate('IFERROR(WEBSERVICE("https://blocked.example"),"blocked")'), 'blocked');
  for (const [formula, code] of [['1/0', '#DIV/0!'], ['SQRT(-1)', '#NUM!'], ['SUM(1e308,1e308)', '#NUM!'], ['ROUND(1)', '#VALUE!'], ['NOT(1,2)', '#VALUE!'], ['SUM("text")', '#VALUE!'], ['HYPERLINK("https://blocked.example")', '#NAME?']]) await assert.rejects(engine.evaluate(formula), error => error instanceof FormulaError && error.message === code, formula);
  await assert.rejects(engine.evaluate('globalThis.fetch("https://blocked.example")'), FormulaError);
  for (const formula of ['[remote.xlsx]Sheet1!A1', '(()=>42)()']) assert.throws(() => parseFormula(formula), FormulaError);
});

test('ranges, criteria, local names and dependency values preserve saved results', async () => {
  const entries = {
    '0:0:0': { value: 'ON' }, '0:1:0': { value: 'On' }, '0:2:0': { value: 'A*B' },
    '0:0:1': { value: 10 }, '0:1:1': { value: 20 }, '0:2:1': { value: 30 },
    '1:0:0': { value: null, formula: 'SUM(Orders!B1:B3)' },
    '1:1:0': { value: 123, formula: '1/0' }
  };
  const engine = evaluator(entries, { definedNames: [{ name: 'Total', value: '999' }, { name: 'Total', value: 'Orders!B1:B3', sheet: 1 }] });
  assert.equal(await engine.cell(1, { row: 0, column: 0 }), 60);
  assert.equal(await engine.cell(1, { row: 1, column: 0 }), 123);
  for (const [formula, expected] of [['SUMIF(Orders!A1:A3,"on",Orders!B1:B3)', 30], ['COUNTIF(Orders!A1:A3,"?n")', 2], ['COUNTIF(Orders!A1:A3,"A~*B")', 1], ['AVERAGEIF(Orders!B1:B3,">=20")', 25], ['SUM(Total)', 60], ['AVERAGE(Orders!B1:B3)', 20]]) assert.equal(await engine.evaluate(formula, 1), expected, formula);
});

test('COUNT/COUNTA handle empty cells, strings, booleans and errors correctly', async () => {
  const engine = evaluator({ '0:0:0': { value: 3 }, '0:1:0': { value: '' }, '0:2:0': { value: false }, '0:3:0': { error: '#DIV/0!' } });
  assert.equal(await engine.evaluate('COUNT(A1:A5)'), 1);
  assert.equal(await engine.evaluate('COUNTA(A1:A5)'), 4);
  await assert.rejects(engine.evaluate('SUM(A1:A5)'), /#DIV\/0!/);
});

test('cycles, oversized expressions, cancellation and operation budgets stop bounded calculations', async () => {
  const cyclic = evaluator({ '0:0:0': { value: null, formula: 'A2' }, '0:1:0': { value: null, formula: 'A1' } });
  await assert.rejects(cyclic.cell(0, { row: 0, column: 0 }), /#REF!/);
  await assert.rejects(evaluator({}, { maxOperations: 10 }).evaluate('SUM(A1:A100)'), /#NUM!/);
  await assert.rejects(evaluator({}, { cancelled: () => true }).evaluate('SUM(1,2)'), error => error.name === 'AbortError');
  assert.throws(() => parseFormula('1+'.repeat(20000)), FormulaError);
});

test('shared formulas rebase relative references, preserve fixed coordinates and never modify string literals', () => {
  assert.equal(rebaseFormula('SUM(A1,$B2,C$3,$D$4,"A1",\'Sheet A1\'!A1,LOG10(A1),Sheet1!A1)', 2, 1), 'SUM(B3,$B4,D$3,$D$4,"A1",\'Sheet A1\'!B3,LOG10(B3),Sheet1!B3)');
  assert.equal(rebaseFormula('A1', -1, 0), '#REF!');
});

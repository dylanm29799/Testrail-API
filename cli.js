const yargs = require('yargs');

// --- CLI options ---
const argv = yargs
  .option('run', {
    alias: 'r',
    describe: 'TestRail run ID',
    type: 'number',
    demandOption: true
  })
  .option('tests', {
    alias: 't',
    describe: 'Tests to include (e.g. 1-10,13,15)',
    type: 'string'
  })
  .option('exclude', {
    alias: 'x',
    describe: 'Tests to exclude (e.g. 2,7)',
    type: 'string'
  })
  .help()
  .argv;

// --- Helper: parse test list like "1-5,7,9" ---
function parseTestInput(input) {
  const testSet = new Set();

  if (!input) return testSet;

  const parts = input.split(',').map(p => p.trim());

  for (const part of parts) {
    if (part.includes('-')) {
      const [start, end] = part.split('-').map(Number);
      for (let i = start; i <= end; i++) {
        testSet.add(i);
      }
    } else {
      const val = Number(part);
      if (!isNaN(val)) testSet.add(val);
    }
  }

  return testSet;
}

// --- Apply logic ---
const ALL_TESTS = Array.from({ length: 100 }, (_, i) => i + 1); // replace with real total if known

const includeSet = argv.tests ? parseTestInput(argv.tests) : new Set(ALL_TESTS);
const excludeSet = argv.exclude ? parseTestInput(argv.exclude) : new Set();

const finalTests = Array.from(includeSet).filter(test => !excludeSet.has(test)).sort((a, b) => a - b);

// --- Output ---
console.log(`\n✅ Run ID: ${argv.run}`);
console.log(`🧪 Tests to run: ${finalTests.join(', ') || '[none]'}`);

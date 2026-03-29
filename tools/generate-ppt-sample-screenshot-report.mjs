import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { createServer } from 'node:http';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);

const root = process.cwd();
const distRoot = path.join(root, 'examples', 'basic', 'dist');
const sampleRoot = process.env.PPT_SAMPLE_ROOT ?? path.join(process.env.HOME ?? '', 'Desktop', 'ppt-samples');
const cacheRoot = path.join(root, '.omx', 'cache', 'playwright-runner');
const screenshotRoot = path.join(root, 'benchmarks', 'reports', 'ppt-sample-screenshots');
const reportPath = path.join(root, 'benchmarks', 'reports', 'ppt-sample-screenshot-report.json');
const port = Number(process.env.PPT_SAMPLE_SCREENSHOT_PORT ?? 4174);

const declaredSlides = {
  sample1: [1, 2, 10],
  sample2: [1, 2, 3],
  sample3: [1, 2],
  sample4: [1, 2],
  sample5: [1, 2],
  sample6: [1, 2]
};

await stat(distRoot);
await mkdir(cacheRoot, { recursive: true });
await mkdir(screenshotRoot, { recursive: true });

async function ensurePlaywright() {
  const packageJsonPath = path.join(cacheRoot, 'package.json');
  try {
    await stat(path.join(cacheRoot, 'node_modules', 'playwright'));
  } catch {
    await writeFile(packageJsonPath, JSON.stringify({ name: 'playwright-runner', private: true }, null, 2));
    await execFileAsync('npm', ['install', 'playwright@1.58.2'], { cwd: cacheRoot });
    await execFileAsync('npx', ['playwright', 'install', 'chromium'], { cwd: cacheRoot });
  }
}

function staticServer(rootDir) {
  return createServer(async (req, res) => {
    const urlPath = req.url === '/' ? '/index.html' : req.url ?? '/index.html';
    const filePath = path.join(rootDir, urlPath.replace(/\?.*$/, ''));
    try {
      const bytes = await readFile(filePath);
      const ext = path.extname(filePath);
      const type = ext === '.html'
        ? 'text/html'
        : ext === '.js'
          ? 'text/javascript'
          : ext === '.map'
            ? 'application/json'
            : 'application/octet-stream';
      res.writeHead(200, { 'content-type': type });
      res.end(bytes);
    } catch {
      res.writeHead(404);
      res.end('not found');
    }
  });
}

function readPngSize(bytes) {
  if (bytes.toString('ascii', 1, 4) !== 'PNG') {
    return null;
  }

  return {
    width: bytes.readUInt32BE(16),
    height: bytes.readUInt32BE(20)
  };
}

await ensurePlaywright();

const server = staticServer(distRoot);
await new Promise((resolve) => server.listen(port, '127.0.0.1', resolve));

try {
  const runnerPath = path.join(cacheRoot, 'capture.mjs');
  await writeFile(runnerPath, `
import { chromium } from 'playwright';
import { mkdir } from 'node:fs/promises';
const [,, baseUrl, samplePath, outputPath, slideIndex] = process.argv;
await mkdir(new URL('file://' + outputPath).pathname.split('/').slice(0, -1).join('/'), { recursive: true });
const browser = await chromium.launch({ headless: true });
const page = await browser.newPage({ viewport: { width: 1500, height: 1300, deviceScaleFactor: 1 } });
await page.goto(baseUrl, { waitUntil: 'networkidle' });
await page.setInputFiles('#file-input', samplePath);
await page.waitForTimeout(1500);
for (let step = 0; step < Number(slideIndex); step += 1) {
  await page.click('#slide-next-button');
  await page.waitForTimeout(250);
}
await page.locator('#preview .ooxml-pptx-slide-canvas').screenshot({ path: outputPath });
await browser.close();
`);

  const results = [];
  for (const [sampleName, slideNumbers] of Object.entries(declaredSlides)) {
    const sampleDir = path.join(sampleRoot, sampleName);
    const pptxPath = path.join(sampleDir, 'sample.pptx');
    for (const slideNumber of slideNumbers) {
      const screenshotPath = path.join(screenshotRoot, sampleName, `sample.${String(slideNumber).padStart(3, '0')}.png`);
      const referencePath = path.join(sampleDir, `sample.${String(slideNumber).padStart(3, '0')}.png`);

      await mkdir(path.dirname(screenshotPath), { recursive: true });
      await execFileAsync(
        'node',
        [runnerPath, `http://127.0.0.1:${port}/`, pptxPath, screenshotPath, String(slideNumber - 1)],
        { cwd: cacheRoot }
      );

      const referenceBytes = await readFile(referencePath);
      const screenshotBytes = await readFile(screenshotPath);
      results.push({
        sample: sampleName,
        slide: slideNumber,
        referencePath: path.relative(root, referencePath).replaceAll('\\', '/'),
        screenshotPath: path.relative(root, screenshotPath).replaceAll('\\', '/'),
        referenceSize: readPngSize(referenceBytes),
        screenshotSize: readPngSize(screenshotBytes)
      });
    }
  }

  await writeFile(reportPath, JSON.stringify({
    generatedAt: new Date().toISOString(),
    sampleRoot,
    screenshotRoot: path.relative(root, screenshotRoot).replaceAll('\\', '/'),
    results
  }, null, 2));
  console.log(JSON.stringify({ reportPath: path.relative(root, reportPath).replaceAll('\\', '/'), count: results.length }, null, 2));
} finally {
  await new Promise((resolve, reject) => server.close((error) => error ? reject(error) : resolve()));
}

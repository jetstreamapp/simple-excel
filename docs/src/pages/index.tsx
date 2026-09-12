import type { ReactNode } from 'react';
import Link from '@docusaurus/Link';
import useDocusaurusContext from '@docusaurus/useDocusaurusContext';
import CodeBlock from '@theme/CodeBlock';
import Layout from '@theme/Layout';

const GITHUB_URL = 'https://github.com/jetstreamapp/simple-excel';

const WRITE_EXAMPLE = `import { collectToBlob, createWorkbookWriter } from '@jetstreamapp/simple-excel';

const sink = collectToBlob();
const workbook = createWorkbookWriter(sink);
const sheet = workbook.addSheet('Accounts', { header: ['Id', 'Name'], freeze: { rows: 1 } });

for await (const record of records) {
  await sheet.writeRow([record.Id, record.Name]);
}

await sheet.close();
await workbook.close();
const blob = await sink.result();`;

function HomepageHeader(): ReactNode {
  const { siteConfig } = useDocusaurusContext();
  return (
    <header className="hero hero--primary">
      <div className="container text--center">
        <h1 className="hero__title">{siteConfig.title}</h1>
        <p className="hero__subtitle">{siteConfig.tagline}</p>
        <div className="margin-top--lg">
          <Link className="button button--secondary button--lg margin-horiz--sm" to="/docs/intro">
            Get started
          </Link>
          <Link className="button button--outline button--secondary button--lg margin-horiz--sm" to={GITHUB_URL}>
            GitHub
          </Link>
        </div>
      </div>
    </header>
  );
}

export default function Home(): ReactNode {
  const { siteConfig } = useDocusaurusContext();
  return (
    <Layout title={siteConfig.title} description="Streaming, dependency-free xlsx reader and writer for browsers and Node.">
      <HomepageHeader />
      <main className="container margin-vert--xl">
        <div className="row">
          <div className="col col--4">
            <h2>Flat memory</h2>
            <p>
              Rows are pushed to a sink as they arrive and pulled through an async iterator on the way back. There is no in-memory
              worksheet, no whole-sheet XML string and no unbounded shared-string table.
            </p>
          </div>
          <div className="col col--4">
            <h2>No dependencies</h2>
            <p>
              The zip container, the XML tokenizer and the SpreadsheetML layer are all in the package, and compression is the
              platform&apos;s own <code>CompressionStream</code>. Nothing to audit but this.
            </p>
          </div>
          <div className="col col--4">
            <h2>Checked against Excel</h2>
            <p>
              Built against a committed corpus of real files and an oracle that opens every written file in Excel, LibreOffice, openpyxl,
              calamine and the Open XML SDK validator.
            </p>
          </div>
        </div>
        <div className="row margin-top--xl">
          <div className="col col--8 col--offset-2">
            <CodeBlock language="ts">{WRITE_EXAMPLE}</CodeBlock>
          </div>
        </div>
      </main>
    </Layout>
  );
}

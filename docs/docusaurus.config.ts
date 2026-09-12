import path from 'path';
import { themes as prismThemes } from 'prism-react-renderer';
import type { Config } from '@docusaurus/types';
import type * as Preset from '@docusaurus/preset-classic';

const config: Config = {
  title: 'simple-excel',
  tagline: 'Streaming xlsx that opens cleanly everywhere',
  favicon: 'img/favicon.ico',

  future: {
    v4: true,
  },

  url: 'https://simple-excel.dev',
  baseUrl: '/',

  organizationName: 'jetstreamapp',
  projectName: 'simple-excel',
  deploymentBranch: 'gh-pages',
  trailingSlash: false,

  onBrokenLinks: 'throw',

  i18n: {
    defaultLocale: 'en',
    locales: ['en'],
  },

  plugins: [
    function aliasPlugin() {
      return {
        name: 'alias-local-package',
        configureWebpack() {
          return {
            resolve: {
              alias: {
                '@jetstreamapp/simple-excel': path.resolve(__dirname, '..', 'dist', 'esm', 'index.mjs'),
              },
            },
          };
        },
      };
    },
    [
      'docusaurus-plugin-llms',
      {
        title: 'simple-excel',
        description:
          'A streaming, dependency-free xlsx reader and writer for browsers and Node, built for large files that must open cleanly in Excel, Google Sheets, LibreOffice and Numbers.',
        generateLLMsTxt: true,
        generateLLMsFullTxt: true,
      },
    ],
  ],

  presets: [
    [
      'classic',
      {
        docs: {
          sidebarPath: './sidebars.ts',
          editUrl: 'https://github.com/jetstreamapp/simple-excel/tree/main/docs/',
        },
        blog: false,
        theme: {
          customCss: './src/css/custom.css',
        },
      } satisfies Preset.Options,
    ],
  ],

  themeConfig: {
    colorMode: {
      defaultMode: 'dark',
      respectPrefersColorScheme: true,
    },
    navbar: {
      title: 'simple-excel',
      items: [
        {
          type: 'docSidebar',
          sidebarId: 'docsSidebar',
          position: 'left',
          label: 'Docs',
        },
        {
          to: '/playground',
          label: 'Playground',
          position: 'left',
        },
        {
          type: 'dropdown',
          label: 'LLM Docs',
          position: 'right',
          items: [
            {
              href: 'https://simple-excel.getjetstream.app/llms.txt',
              label: 'llms.txt (summary)',
            },
            {
              href: 'https://simple-excel.getjetstream.app/llms-full.txt',
              label: 'llms-full.txt (complete)',
            },
          ],
        },
        {
          href: 'https://github.com/jetstreamapp/simple-excel',
          label: 'GitHub',
          position: 'right',
        },
        {
          href: 'https://www.npmjs.com/package/simple-excel',
          label: 'npm',
          position: 'right',
        },
      ],
    },
    footer: {
      style: 'dark',
      links: [
        {
          title: 'Docs',
          items: [
            {
              label: 'Getting Started',
              to: '/docs/getting-started',
            },
            {
              label: 'API Reference',
              to: '/docs/api-reference',
            },
            {
              label: 'Functions',
              to: '/docs/functions/logical',
            },
          ],
        },
        {
          title: 'Tools',
          items: [
            {
              label: 'Playground',
              to: '/playground',
            },
            {
              label: 'LLM Docs (llms.txt)',
              href: 'https://simple-excel.getjetstream.app/llms.txt',
            },
            {
              label: 'LLM Docs (full)',
              href: 'https://simple-excel.getjetstream.app/llms-full.txt',
            },
          ],
        },
        {
          title: 'More',
          items: [
            {
              label: 'GitHub',
              href: 'https://github.com/jetstreamapp/simple-excel',
            },
            {
              label: 'npm',
              href: 'https://www.npmjs.com/package/simple-excel',
            },
          ],
        },
      ],
      copyright: `Copyright © ${new Date().getFullYear()} simple-excel. Built with Docusaurus.`,
    },
    prism: {
      theme: prismThemes.github,
      darkTheme: prismThemes.dracula,
      additionalLanguages: ['bash', 'json'],
    },
  } satisfies Preset.ThemeConfig,
};

export default config;

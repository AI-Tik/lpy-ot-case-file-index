import type { Metadata, Viewport } from 'next';
import './globals.css';

const publicSiteUrl = process.env.NEXT_PUBLIC_SITE_URL ?? 'http://localhost:3000';
const publicBasePath = process.env.NEXT_PUBLIC_BASE_PATH ?? '';
const publicAsset = (path: string) => `${publicBasePath}${path}`;

export const metadata: Metadata = {
  metadataBase: new URL(publicSiteUrl),
  title: 'LPY 職業治療個案檔案索引',
  description: '在手機上快速查找、加入、修改及刪除個案檔案位置。',
  applicationName: 'LPY 個案檔案索引',
  manifest: publicAsset('/manifest.webmanifest'),
  icons: {
    icon: [
      { url: publicAsset('/icon-192.png'), sizes: '192x192', type: 'image/png' },
      { url: publicAsset('/icon-512.png'), sizes: '512x512', type: 'image/png' },
    ],
    apple: [{ url: publicAsset('/apple-touch-icon.png'), sizes: '180x180', type: 'image/png' }],
  },
  appleWebApp: { capable: true, statusBarStyle: 'default', title: '個案索引' },
  formatDetection: { telephone: false },
  openGraph: {
    title: 'LPY 職業治療個案檔案索引',
    description: '快速查找檔案顏色及編號',
    images: [{ url: publicAsset('/og.jpg'), width: 1200, height: 800 }],
  },
  twitter: {
    card: 'summary_large_image',
    title: 'LPY 職業治療個案檔案索引',
    description: '快速查找檔案顏色及編號',
    images: [publicAsset('/og.jpg')],
  },
};

export const viewport: Viewport = {
  width: 'device-width', initialScale: 1, maximumScale: 1, themeColor: '#247b6d',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-Hant"><body>{children}</body></html>;
}

import type { NextConfig } from "next";
const isGithubPages = process.env.GITHUB_ACTIONS === "true";
const nextConfig: NextConfig = {
  reactStrictMode: true,
  output: "export",
  images: { unoptimized: true },
  ...(isGithubPages ? { basePath: "/kit", assetPrefix: "/kit/" } : {}),
};
export default nextConfig;

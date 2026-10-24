import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Hide the Next.js "N" / Route Info badge in local development.
  // Compile and runtime errors still surface.
  devIndicators: false,
};

export default nextConfig;

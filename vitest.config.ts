import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    globals: true,
    environment: "node",
    include: ["src/**/*.test.ts"],
    exclude: ["node_modules", "dist", "src/commands/**"],
    coverage: {
      provider: "v8",
      include: ["src/**/*.ts"],
      exclude: ["src/commands/**", "src/**/*.test.ts"],
    },
  },
});

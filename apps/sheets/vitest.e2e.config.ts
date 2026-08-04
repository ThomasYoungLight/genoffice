import { defineConfig } from 'vitest/config'

/**
 * End-to-end runs, kept out of `npm test` on purpose.
 *
 * They need a built app (`npm run build` plus the Rust sidecar), they launch a
 * real Electron process per file, and they take minutes rather than seconds.
 * They also contribute nothing to coverage — the vitest v8 provider cannot
 * instrument a separate process — so mixing them into the unit run would make
 * both numbers harder to read.
 */
export default defineConfig({
  test: {
    environment: 'node',
    include: ['tests-e2e/**/*.e2e.ts'],
    // One Electron at a time: several instances contend for the sidecar and
    // for the window server, and a flake there looks like a product bug.
    fileParallelism: false,
    pool: 'forks',
    testTimeout: 180_000,
    hookTimeout: 180_000,
    teardownTimeout: 30_000,
    retry: 0,
  },
})

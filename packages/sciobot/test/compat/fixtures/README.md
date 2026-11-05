The native gallery fixtures record an actual read and write through Runtime G at `c3bf6048a`.
That runtime predates payload version 3. Its private writer refuses all four opaque gallery records.
Its core writer preserves the native images, captions, relationships, and payload data.
The JSON file records the bundle hashes and both document hashes.

Run the reproduction from the repository root with an immutable Runtime G snapshot:

```bash
bun --preload ./scripts/test/preload-dom.ts \
  packages/sciobot/test/compat/gallery-v3-old-writer.ts \
  /path/to/runtime-g c3bf6048a /tmp/native-gallery-compat
```

The regression compares the recorded input and output through the current reader.
The reproduction script executes the old reader and writer before it creates those files.

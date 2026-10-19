import sys, zipfile, os
src, out = sys.argv[1], sys.argv[2]
files = []
for root, _, fs in os.walk(src):
    for f in fs:
        full = os.path.join(root, f); files.append(os.path.relpath(full, src))
files.sort(key=lambda p: (p != "[Content_Types].xml", not p.startswith("_rels"), p))
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for f in files: z.write(os.path.join(src, f), f)

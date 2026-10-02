from pathlib import Path
import base64, zlib
parts = [
    'scripts/ppt10021021_chunk1.txt',
    'scripts/ppt10021021_chunk2.txt',
    'scripts/ppt10021021_chunk3_fixed.txt',
    'scripts/ppt10021021_chunk4.txt',
]
payload = ''.join(Path(p).read_text(encoding='utf-8') for p in parts)
source = zlib.decompress(base64.b64decode(payload)).decode('utf-8')
exec(compile(source, 'build_refined_ppt_repo.py', 'exec'))

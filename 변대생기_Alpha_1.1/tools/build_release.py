"""Explicit source-only release; never collect user documents."""
from pathlib import Path
import zipfile
ROOT = Path(__file__).resolve().parents[1]
FILES = ['.gitignore', 'app.py', 'generator.py', 'requirements.txt', 'README.md', 'CHANGELOG.md', 'VALIDATION.md',
         'INSTALL.bat', 'RUN.bat', 'TEST.bat', 'DIAGNOSE.bat', 'UPLOAD_TO_GITHUB.bat', 'UPLOAD_TO_GITHUB.ps1',
         'tests/test_generator.py', 'tests/test_matching_review.py', 'tools/build_release.py']

def main():
    names = FILES
    for name in names:
        if not (ROOT / name).is_file(): raise FileNotFoundError(name)
    target = ROOT / 'dist' / 'Byeondaesaenggi_Alpha_1.1.zip'; target.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as z:
        for name in names: z.write(ROOT / name, '변대생기_Alpha_1.1/' + name)
    with zipfile.ZipFile(target) as z:
        assert z.testzip() is None
        assert len(z.namelist()) == len(names)
    print(target)

if __name__ == '__main__': main()


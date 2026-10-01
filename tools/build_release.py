"""Build a source-only distribution; no input documents or virtual environment."""
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
FILES = [
    '.gitignore', 'analyzer.py', 'app.py', 'pair_analysis.py', 'pair_runner.py', 'word_capture.py',
    'requirements.txt', 'README.md', 'CHANGELOG.md', 'VALIDATION.md',
    'INSTALL.bat', 'RUN.bat', 'DIAGNOSE.bat', 'TEST.bat',
    'UPLOAD_TO_GITHUB.bat', 'UPLOAD_TO_GITHUB.ps1',
    'tests/test_pair_analysis.py', 'tests/test_alpha04.py',
    'tools/validate_sample.py', 'tools/build_release.py', 'AT_GENERATOR_history.bundle',
]


def main():
    for name in FILES:
        if not (ROOT/name).is_file():
            raise FileNotFoundError(name)
    target = ROOT/'dist'/'Byeonbungi_Alpha_0.4.zip'
    target.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name in FILES:
            archive.write(ROOT/name, '변분기_Alpha_0.4/'+name)
    with zipfile.ZipFile(target) as archive:
        assert archive.testzip() is None
        assert len(archive.namelist()) == len(FILES)
    print(target)


if __name__ == '__main__':
    main()

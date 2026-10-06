"""Build the two-folder Alpha 1.1 distribution from explicit source lists."""
from pathlib import Path
import hashlib
import runpy
import zipfile

SUITE = Path(__file__).resolve().parents[2]
FOLDERS = ('변분기_Alpha_1.1', '변대생기_Alpha_1.1')
RELEASE_DATE = (2026, 10, 2, 0, 0, 0)


def source_files():
    sources = []
    for folder in FOLDERS:
        root = SUITE / folder
        config = runpy.run_path(str(root / 'tools' / 'build_release.py'))
        for name in config['FILES']:
            relative = Path(name)
            if relative.is_absolute() or '..' in relative.parts:
                raise ValueError('Release source must be inside its version folder: ' + name)
            source = root / relative
            if not source.is_file():
                raise FileNotFoundError(source)
            source.resolve().relative_to(root.resolve())
            sources.append((source, folder + '/' + relative.as_posix()))
    if len({name for _, name in sources}) != len(sources):
        raise ValueError('Duplicate release source paths')
    return sources


def main():
    sources = source_files()
    target = SUITE / 'AT_GENERATOR_Alpha_1.1.zip'
    with zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for source, name in sources:
            info = zipfile.ZipInfo(name, date_time=RELEASE_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o100644 << 16
            archive.writestr(info, source.read_bytes(), compresslevel=9)
    with zipfile.ZipFile(target) as archive:
        if archive.testzip() is not None:
            raise ValueError('Invalid release ZIP')
        if archive.namelist() != [name for _, name in sources]:
            raise ValueError('Release source list differs')
        if {name.split('/')[0] for name in archive.namelist()} != set(FOLDERS):
            raise ValueError('Release must contain exactly two version folders')
        for source, name in sources:
            if archive.read(name) != source.read_bytes():
                raise ValueError('Release bytes differ: ' + name)
    print(target)
    print(f'{len(sources)} source files; two version folders')
    print('SHA256 ' + hashlib.sha256(target.read_bytes()).hexdigest())


if __name__ == '__main__':
    main()

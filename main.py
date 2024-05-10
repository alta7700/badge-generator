import argparse
from pathlib import Path
from typing import Iterable

import photoshop.api as ps


app = ps.Application()
doc = app.activeDocument


lines_layers = {
    21: doc.artLayers.getByName(f'__ФИО21'),
    22: doc.artLayers.getByName(f'__ФИО22'),
    31: doc.artLayers.getByName(f'__ФИО31'),
    32: doc.artLayers.getByName(f'__ФИО32'),
    33: doc.artLayers.getByName(f'__ФИО33'),
    41: doc.artLayers.getByName(f'__ФИО41'),
    42: doc.artLayers.getByName(f'__ФИО42'),
    43: doc.artLayers.getByName(f'__ФИО43'),
    44: doc.artLayers.getByName(f'__ФИО44'),
}


def set_visible(lines: Iterable[int]):
    all_lines = set(lines_layers.keys())
    for line in lines:
        lines_layers[line].visible = True
        all_lines.remove(line)
    for line in all_lines:
        lines_layers[line].visible = False


def fill_fio_layers(fio_lines: list[str]):
    match len(fio_lines):
        case 1:
            lines = (32,)
        case 2:
            lines = (21, 22)
        case 3:
            lines = (31, 32, 33)
        case 4:
            lines = (41, 42, 43, 44)
        case _:
            raise
    set_visible(lines)
    for i, line in enumerate(lines):
        lines_layers[line].textItem.contents = fio_lines[i]


def main(lines: list[int], results_dir: Path, csv: Path, id_mode: bool):
    results_dir = results_dir.absolute()
    with open(csv, 'r', encoding='utf-8') as studs:
        studs.readline()  # read header
        for stud in studs:
            if not stud:
                continue
            if id_mode:
                user_id, course, group, *fio_lines = (x for x in stud.strip('\n').split(';') if x)
                out_filename = f'{user_id}.png'
            else:
                course, group, *fio_lines = (x for x in stud.strip('\n').split(';') if x)
                out_filename = f'{str(group).zfill(3)}гр_{" ".join(fio_lines)}.png'
            if len(fio_lines) not in lines:
                continue
            for layer in doc.artLayers:
                if layer.kind == ps.LayerKind.TextLayer:
                    match layer.name:
                        case '__КУРС':
                            layer.textItem.contents = course
                        case '__ГРУППА':
                            layer.textItem.contents = group
            fill_fio_layers(fio_lines)

            png_file = results_dir / out_filename
            doc.saveAs(png_file.as_posix(), ps.PNGSaveOptions(), asCopy=True)
            print(png_file.name)


if __name__ == '__main__':
    BASE_PATH = Path(__file__).parent

    parser = argparse.ArgumentParser()
    parser.add_argument('--csv', type=Path, nargs='?', default=BASE_PATH / 'studs.csv')
    parser.add_argument('--results-dir', type=Path, nargs='?', default=BASE_PATH / 'results')
    parser.add_argument('--lines', type=int, nargs='*', default=[1, 2, 3, 4])
    parser.add_argument('--id-mode', action='store_true', default=False)
    options = parser.parse_args().__dict__
    rd = options['results_dir']
    if not rd.exists():
        rd.mkdir()
    assert rd.is_dir()
    main(**options)

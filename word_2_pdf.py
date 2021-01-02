import comtypes.client
from pathlib import Path
from htutil import file


def get_powerpoint():
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1
    return powerpoint


def ppt_to_pdf(input_file_path: Path, powerpoint: None):
    output_file_path = input_file_path.parents[1] / \
        'output' / f'{input_file_path.stem}.pdf'
    print(str(output_file_path.absolute()))

    powerpoint_is_None = powerpoint == None

    if powerpoint_is_None:
        powerpoint = get_powerpoint()

    deck = powerpoint.Presentations.Open(str(input_file_path.absolute()))
    formatType = 32  # formatType = 32 for ppt to pdf
    deck.SaveAs(str(output_file_path.absolute()), formatType)
    deck.Close()

    if powerpoint_is_None:
        powerpoint.Quit()


def main():
    file.create_dir_if_not_exist('input')
    file.create_dir_if_not_exist('output')

    dir_source = Path('input')

    powerpoint = get_powerpoint()

    for input_file_path in dir_source.glob('*.ppt*'):
        ppt_to_pdf(input_file_path, powerpoint)

    powerpoint.Quit()


if __name__ == '__main__':
    main()

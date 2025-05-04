import argparse
from utils.ppt_creator import create_presentation

def main():
    parser = argparse.ArgumentParser(description="Convert Excel data to PowerPoint presentation")
    parser.add_argument('--input', type=str, required=True, help='Path to Excel file')
    parser.add_argument('--output', type=str, required=True, help='Path to save PowerPoint file')
    parser.add_argument('--sheet', type=str, default=None, help='Sheet name (optional)')
    
    args = parser.parse_args()

    create_presentation(args.input, args.output, args.sheet)

if __name__ == "__main__":
    main()

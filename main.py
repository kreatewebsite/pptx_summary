import os
import argparse as arg_parse
from pptx_utils import extract_slides_from_pptx, find_and_read_ppt_files
from summarizer import generate_summary


def main(root_dir,output):
    pptx_files = find_and_read_ppt_files(root_dir)

    for pptx_path in pptx_files:
        slides_text = extract_slides_from_pptx(pptx_path)

        base_filename = os.path.basename(pptx_path).split(".")[0]
        summary_dir = os.path.join(output, f"{base_filename}_summary")
        os.makedirs(summary_dir, exist_ok=True)

        for i, slide_text in enumerate(slides_text):
            summary = generate_summary(slide_text)

            summary_filename = f"{base_filename}_summary_slide_{i + 1}.comment"
            summary_path = os.path.join(summary_dir, summary_filename)

            if os.path.exists(summary_path):
                os.remove(summary_path)

            with open(summary_path, 'w', encoding='utf-8') as summary_file:
                summary_file.write(summary)

            print(f"Summary saved for slide {i + 1} in {pptx_path} at {summary_path}")

def single_pptx(pptx_path,output):
    slides_text = extract_slides_from_pptx(pptx_path)

    base_filename = os.path.basename(pptx_path).split(".")[0]
    summary_dir = os.path.join(output, f"{base_filename}_summary")
    os.makedirs(summary_dir, exist_ok=True)

    for i, slide_text in enumerate(slides_text):
        summary = generate_summary(slide_text)

        summary_filename = f"{base_filename}_summary_slide_{i + 1}.comment"
        summary_path = os.path.join(summary_dir, summary_filename)

        if os.path.exists(summary_path):
            os.remove(summary_path)

        with open(summary_path, 'w', encoding='utf-8') as summary_file:
            summary_file.write(summary)

        print(f"Summary saved for slide {i + 1} in {pptx_path} at {summary_path}")
        


def get_args():

    parser = arg_parse.ArgumentParser()

    parser.add_argument("-file", type=str, help="Path to a single text file")
    parser.add_argument("-folder", type=str, help="Path to a folder of txt files")
    parser.add_argument("-output", type=str, help="Path where output file will be saved",required=True)

    args, _ = parser.parse_known_args()
    return args


if __name__ == "__main__":
    args = get_args()
    # print("\nProvide Output path always, even if it doesn't exist\n")

    try:
        if args.file:
            single_pptx(args.file,args.output)
        elif args.folder:
            main(args.folder,args.output)

    except Exception as e:

        print("\n",e)


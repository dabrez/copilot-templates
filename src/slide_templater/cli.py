import argparse
import sys
from .templater import replace_text_in_presentation, load_config
from .analyzer import analyze_presentation
from .llm_analyzer import analyze_presentation_llm

def main():
    parser = argparse.ArgumentParser(description="Slide Templater - Tools for automating PowerPoint presentations.")
    subparsers = parser.add_subparsers(dest="command", help="Subcommand to run")

    # Analyze subcommand
    analyze_parser = subparsers.add_parser("analyze", help="Extract potential variables from a presentation.")
    analyze_parser.add_argument("--input", required=True, help="Path to the source .pptx file.")
    analyze_parser.add_argument("--output", required=True, help="Path to save the generated JSON config file.")
    analyze_parser.add_argument("--method", choices=["nlp", "llm"], default="nlp", help="Method to use for analysis (nlp or llm). Default is nlp.")

    # Template subcommand
    template_parser = subparsers.add_parser("template", help="Replace variables in a presentation.")
    template_parser.add_argument("--input", required=True, help="Path to the source .pptx file.")
    template_parser.add_argument("--config", required=True, help="Path to the JSON configuration file.")
    template_parser.add_argument("--output", required=True, help="Path to save the modified .pptx file.")

    args = parser.parse_args()

    try:
        if args.command == "analyze":
            if args.method == "llm":
                analyze_presentation_llm(args.input, args.output)
            else:
                analyze_presentation(args.input, args.output)
            print(f"Successfully analyzed presentation using {args.method.upper()}. Config saved to: {args.output}")
        elif args.command == "template":
            config = load_config(args.config)
            replace_text_in_presentation(args.input, config, args.output)
            print(f"Successfully processed presentation. Saved to: {args.output}")
        else:
            parser.print_help()
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()

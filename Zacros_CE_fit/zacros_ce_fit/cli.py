# zacros_ce_fit/cli.py
import argparse
from .analysis import run_analysis

def main():
    parser = argparse.ArgumentParser(description="Zacros CE Fit Analysis")
    parser.add_argument("file", help="Path to the AmcBvec Excel file")
    parser.add_argument("--correlation", action="store_true", help="Generate correlation matrices")
    parser.add_argument("--histograms", action="store_true", help="Generate frequency distributions")
    parser.add_argument("--stats", action="store_true", help="Generate statistical summary")
    parser.add_argument("--ncols", type=int, default=None, help="Number of first n columns to include in correlation analysis")
    args = parser.parse_args()

    # If none of the flags are set, do everything
    do_correlation = args.correlation or not (args.correlation or args.histograms or args.stats)
    do_frequency = args.histograms or not (args.correlation or args.histograms or args.stats)
    do_stats = args.stats or not (args.correlation or args.histograms or args.stats)

    run_analysis(
        file_path=args.file,
        n_cols=args.ncols,
        do_correlation=do_correlation,
        do_frequency=do_frequency,
        do_stats=do_stats
    )

if __name__ == "__main__":
    main()
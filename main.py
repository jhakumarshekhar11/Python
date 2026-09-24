"""
Entry point. Anything you save into the `output/` folder here
will automatically be uploaded as a downloadable artifact by the
GitHub Actions workflow (.github/workflows/run.yml) on every commit.

Replace this demo logic with your real code.
"""

import os
import numpy as np
import matplotlib.pyplot as plt

OUTPUT_DIR = "output"


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # --- demo: generate some data and save a chart + a csv ---
    x = np.linspace(0, 10, 100)
    y = np.sin(x)

    plt.figure()
    plt.plot(x, y)
    plt.title("Demo output")
    plt.savefig(os.path.join(OUTPUT_DIR, "demo_plot.png"))

    np.savetxt(os.path.join(OUTPUT_DIR, "demo_data.csv"), np.column_stack([x, y]),
               delimiter=",", header="x,y", comments="")

    print("Done. Files written to ./output/")


if __name__ == "__main__":
    main()

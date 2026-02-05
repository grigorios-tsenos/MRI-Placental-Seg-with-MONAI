# MRI Placental Segmentation with MONAI

This repository contains the working material for my NTUA Diploma Thesis on **placenta segmentation in 3D MRI** using deep learning.

The project is currently thesis-centric: experiments are kept as notebooks and the main deliverable is the LaTeX thesis document.

## Current Thesis Scope

The thesis compares multiple 3D segmentation families under a common MONAI/PyTorch pipeline:

- U-Net
- Attention U-Net
- DynUNet
- UNETR
- SwinUNETR
- SegResNet (two configurations)
- SegMamba (two configurations)

Quantitative and qualitative comparisons are documented inside the thesis sources.

## Repository Layout

- `Thesis Doc/`: LaTeX source, bibliography, figures, and compiled thesis PDF.
- `FINAL RUNS/`: final experiment notebooks used for model training/evaluation runs.
- `UNET Big No of Channels.ipynb`: exploratory notebook from earlier experimentation.
- `Images and data for diploma/`: supporting assets used during writing.

## Building the Thesis PDF

From the project root:

```bash
cd "Thesis Doc"
latexmk -pdf main.tex
```

If `latexmk` is unavailable, you can use `pdflatex` + `bibtex` manually.

## Notes

- This repo no longer uses a single training entry script (such as `swin_822.py`).
- Training logic is maintained in the notebooks under `FINAL RUNS/` and related thesis folders.

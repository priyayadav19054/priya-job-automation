from pathlib import Path
import subprocess
def compile_resume(tex_file, output_dir="pdfs"):
    tex_file = Path(tex_file); output_dir = Path(output_dir); output_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(["pdflatex", "-interaction=nonstopmode", "-halt-on-error", "-output-directory", str(output_dir), str(tex_file)], check=True)
    pdf = output_dir / f"{tex_file.stem}.pdf"
    if not pdf.exists(): raise RuntimeError(f"PDF not found: {pdf}")
    return pdf

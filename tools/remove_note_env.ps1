$ErrorActionPreference = "Stop"

Remove-Item Env:CONDA_PREFIX -ErrorAction SilentlyContinue
Remove-Item Env:CONDA_DEFAULT_ENV -ErrorAction SilentlyContinue
Remove-Item Env:CONDA_PROMPT_MODIFIER -ErrorAction SilentlyContinue
Remove-Item Env:CONDA_SHLVL -ErrorAction SilentlyContinue
Remove-Item Env:CONDA_EXE -ErrorAction SilentlyContinue
Remove-Item Env:_CE_CONDA -ErrorAction SilentlyContinue
Remove-Item Env:_CE_M -ErrorAction SilentlyContinue

conda env remove -n note -y
conda info --envs


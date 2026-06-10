Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$src = Join-Path $root 'src'
$dist = Join-Path $root 'dist'
$build = Join-Path $root 'build'
$release = Join-Path $root 'release'
$appName = 'ccxd_us_reports_tool'
$excludeModules = @(
  'IPython',
  'altair',
  'catboost',
  'dask',
  'datasets',
  'deepspeed',
  'django',
  'distributed',
  'faiss',
  'fastapi',
  'flask',
  'folium',
  'gradio',
  'imageio',
  'jupyterlab',
  'langchain',
  'langchain_community',
  'lightning',
  'matplotlib',
  'modelscope',
  'nbconvert',
  'nbformat',
  'nltk',
  'notebook',
  'bokeh',
  'panel',
  'playwright',
  'plotly',
  'pygame',
  'pyproj',
  'pytest',
  'scipy',
  'skimage',
  'sqlalchemy',
  'statsmodels',
  'tensorflow',
  'timm',
  'torch',
  'torchvision',
  'transformers',
  'xformers',
  'yt_dlp'
)

New-Item -ItemType Directory -Force -Path $dist | Out-Null
New-Item -ItemType Directory -Force -Path $build | Out-Null
New-Item -ItemType Directory -Force -Path $release | Out-Null

Get-ChildItem -Force $dist -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -Force $build -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -Force $release -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem -Force $root -Filter '*.spec' -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
if (Test-Path (Join-Path $root 'test_outputs')) {
  Remove-Item -LiteralPath (Join-Path $root 'test_outputs') -Recurse -Force
}

$args = @(
  '-m', 'PyInstaller',
  '--noconfirm',
  '--clean',
  '--windowed',
  '--name', $appName,
  '--distpath', $dist,
  '--workpath', $build,
  '--paths', $src,
  '--collect-submodules', 'ccxd_us_reports_app',
  '--collect-all', 'openpyxl',
  '--collect-all', 'babel',
  '--collect-all', 'lxml'
)
foreach ($module in $excludeModules) {
  $args += @('--exclude-module', $module)
}
$args += (Join-Path $src 'ccxd_us_reports_app\\app.py')

python @args
python (Join-Path $root 'package_release.py')

RuntimeError: This app has encountered an error. The original error message is redacted to prevent data leaks. Full error details have been recorded in the logs (if you're on Streamlit Cloud, click on 'Manage app' in the lower right of your app).
Traceback:

File "/mount/src/releve-to-excell/app.py", line 8, in <module>
    from paddleocr import PaddleOCR
File "/home/adminuser/venv/lib/python3.14/site-packages/paddleocr/__init__.py", line 17, in <module>
    from ._models import (
    ...<13 lines>...
    )
File "/home/adminuser/venv/lib/python3.14/site-packages/paddleocr/_models/__init__.py", line 15, in <module>
    from .chart_parsing import ChartParsing
File "/home/adminuser/venv/lib/python3.14/site-packages/paddleocr/_models/chart_parsing.py", line 16, in <module>
    from ._doc_vlm import (
    ...<2 lines>...
    )
File "/home/adminuser/venv/lib/python3.14/site-packages/paddleocr/_models/_doc_vlm.py", line 21, in <module>
    from .base import PaddleXPredictorWrapper, PredictorCLISubcommandExecutor
File "/home/adminuser/venv/lib/python3.14/site-packages/paddleocr/_models/base.py", line 17, in <module>
    from paddlex import create_predictor
File "/home/adminuser/venv/lib/python3.14/site-packages/paddlex/__init__.py", line 45, in <module>
    _initialize()
    ~~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/paddlex/__init__.py", line 42, in _initialize
    repo_manager.initialize()
    ~~~~~~~~~~~~~~~~~~~~~~~^^
File "/home/adminuser/venv/lib/python3.14/site-packages/paddlex/repo_manager/core.py", line 214, in initialize
    raise RuntimeError(
        "PDX has already been initialized. Reinitialization is not supported."
    )

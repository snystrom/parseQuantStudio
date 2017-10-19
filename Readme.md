# parseQuantStudio.py
## Author: Spencer Nystrom

Takes excel file output & plots CT, amplification curve, & melt-curve data.

Also provides library (using import) for working interactively with QuantStudio output files 

### Importing as library:
For use inside another python script / jupyter notebook add the following to your header:

```{python}
import sys
sys.path.append('path/to/parseQuantStudio')
from parseQuantStudio import *
```

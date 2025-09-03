import os

EXPERIMENT_NAMES = "AS003_532nm"
data_dir = os.path.join('/Users/souchaud/Documents/Travail/CitizenSers/Spectroscopie/',EXPERIMENT_NAMES)


spectrocopy_files = [f for f in os.listdir(data_dir)
                    if f.endswith('.txt')]

print(spectrocopy_files)

import panda as pd

pd.read_txt(os.path.join(data_dir,spectrocopy_files[0]), skiprows=2, delimiter='\t')
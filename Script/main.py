import pandas as pd
import pathlib
import os, sys

import matplotlib
matplotlib.use('Agg')

import matplotlib.pyplot as plt

# path_to_file = input("Please, insert the folder path where is located the file you want to Analise:")
filename = input("Please, insert the name of the file you want to Analise(with Extension):")

# GET THE PATH OF THE FOLDER WHERE THE FILE AND EXECUTABLE MUST BE RUNNING THE SCRIPT VERSION
path_to_file = pathlib.Path(__file__).parent.resolve().joinpath(filename)

# GET THE PATH OF THE FOLDER WHERE THE FILE AND EXECUTABLE MUST BE THE .EXE VERSION
# path_to_file = os.path.dirname(sys.executable)
# path_to_file = os.path.join(path_to_file, filename)


print('Collecting Data...\n')

excel_data = pd.read_excel(path_to_file)

# LOAD THE FILE
data = pd.DataFrame(excel_data, columns=['TestVoltage1V8Generation','TestVbat','TestRadioRx',
                                         'TestRadioTx','TestTemperature','TestLedCurrent',
                                         'TestRadioRxCurrent','TestRadioTxCurrent','TestAccelerometerGravityAverage',
                                         'TestAccelerometerGravityAverage_2','TestAccelerometerGravityAverage_3',
                                         'TestAccelerometerGravityAverage_4','TestAccelerometerGravityAverage_5',
                                         'TestAccelerometerGravityAverage_6','TestRtc','TestShelfModeCurrentConsumption',
                                         'TestDeepSleepCurrentConsumption','TestCapsDischargeTime'])



# pd.set_option('display.max_rows', data.shape[0]+1)
print("The content of the file is:\n", data)



fig, axes = plt.subplots(6, 3, figsize=(25, 35))
axe = axes.ravel()

# CREATE ONE SUBPLOT FOR EACH RESULT COLUMN WITH THE RESULTS
for i, c in enumerate(data.columns):
    # data[c] = data[c].astype(float)
    data[c] = data[c].apply(pd.to_numeric, errors='coerce').astype(float).dropna()

    ax = data[c].hist(ax=axe[i],bins=25)
    axe[i].title.set_text(c)
    ax.set_xticks(axe[i].get_xticks(), rotation=180)

    # PUT THE NUMBER OF COUNTS ABOVE THE RECTANGLES IN THE HISTOGRAMS FOR EACH SUBPLOT
    for rect in ax.patches:
        y_value = rect.get_height()
        x_value = rect.get_x() + rect.get_width() / 2
        space = 1
        label = "{:.0f}".format(y_value)
        if label != '0':
            ax.annotate(label, (x_value, y_value), xytext=(0, space), textcoords="offset points", ha='center', va='bottom', fontsize=10)

# PUT THE TITLE ON THE START OF THE FILES (SAME TITLE OF THE .xlsx FILE)
fig.suptitle("Test Report for " + filename[0:-5], fontsize=50, fontweight="bold", family="sans-serif")


plt.subplots_adjust(left=0.04,
                    bottom=0.05,
                    right=0.96,
                    top=0.935,
                    wspace=0.13,
                    hspace=0.2)


plt.margins(0.1)
# plt.show()

print("Generating png file...")
plt.savefig(filename[0:-5]+'.png')
print("Generating pdf file")
plt.savefig(filename[0:-5]+'.pdf')

import openpyxl
import numpy as np
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures

def PolynomialRegression():
    wb = openpyxl.load_workbook(r'c:dataset.xlsx', data_only=True)

    training_data = []
    sheet1 = wb['Normailised']
    sheet2 = wb['Predictions']
    day = sheet1['G']
    time = sheet1['H']
    activity = [cell.value for cell in sheet1['I'][0:]]
    for i, j in zip(day, time):
        training_data.append([i.value, j.value])
    x = np.array(training_data) 
    y = np.array(activity)
    #print(training_data)
    x_ = PolynomialFeatures(degree=2, include_bias=False).fit_transform(x)
    pre_day = sheet2['B']
    pre_time = sheet2['C']
    predictions = []
    for i, j in zip(pre_day, pre_time):
        predictions.append([i.value, j.value])
    input = np.array(predictions)
    input_= PolynomialFeatures(degree=2, include_bias=False).fit_transform(input)

    model = LinearRegression().fit(x_, y)
    r_sq = model.score(x_, y)
    intercept, coefficients = model.intercept_, model.coef_
    y_pred = model.predict(input_)
    print(f"coefficient of determination: {r_sq}")
    print(f"intercept: {intercept}")
    print(f"coefficients:\n{coefficients}")
    print(f"predicted response:\n{y_pred}")
    count = 1
    for i in y_pred:
        sheet2.cell(count, column=4).value = i
        count+=1
    wb.save('dataset.xlsx')


if __name__ == '__main__':
    PolynomialRegression()
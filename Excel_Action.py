import pandas as pd


# First, we need to get the path of file execution
# 获取文件路径并获取文件内容
def Excel_action(Path, PathFile, ticknum=500, newfilename="NewSheet"):
    getPath = Path  # 获取路径，最后生成文件使用
    df = pd.read_excel(PathFile)
    lineNum = df.count()

    # 获取内容框架
    # frame = pd.DataFrame(df)
    # print("获取到所有的值:\n{0}".format(df))  # 格式化输出
    # print(lineNum['姓名'])

    # Second, we need to get order by the time
    orderByTimeFrame = df.sort_values(by='报名时间', ascending=True)

    # 增加第一列ID
    orderByTimeFrame.insert(0, "ID", [x for x in range(1, lineNum['姓名'] + 1, 1)])

    # choose the 500%==1 'line'
    TakeTheSpeciaLine = (orderByTimeFrame['ID'] % ticknum == 1)

    TheFinishedTake = orderByTimeFrame[TakeTheSpeciaLine]

    newFileSetPath = getPath + "\\" + newfilename + ".xls"

    TheFinishedTake.to_excel(newFileSetPath)

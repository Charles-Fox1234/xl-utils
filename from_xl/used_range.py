from xlsxwriter.utility import xl_col_to_name
import xlwings as xw
import pandas as pd


def used_range(sht):
    row = last_row(sht)
    column,col_letter = last_column(sht)
    
    return "a1:"+col_letter+str(row),row,column
    
def last_row(sht):
    row_cell = sht.api.Cells.Find(What="*",
                   After=sht.api.Cells(1, 1),
                   LookAt=xw.constants.LookAt.xlPart,
                   LookIn=xw.constants.FindLookIn.xlFormulas,
                   SearchOrder=xw.constants.SearchOrder.xlByRows,
                SearchDirection=xw.constants.SearchDirection.xlPrevious,
                       MatchCase=False)
    
    return row_cell.Row

def last_column(sht):
    column_cell = sht.api.Cells.Find(What="*",
                      After=sht.api.Cells(1, 1),
                      LookAt=xw.constants.LookAt.xlPart,
                      LookIn=xw.constants.FindLookIn.xlFormulas,
                      SearchOrder=xw.constants.SearchOrder.xlByColumns,
                      SearchDirection=xw.constants.SearchDirection.xlPrevious,
                      MatchCase=False)
    
    c = column_cell.Column
    return c, xl_col_to_name(c-1)

def to_df(sht,n_index_cols = 1):
    rng = used_range(sht)[0]
    data = sht.range(rng).value
    if n_index_cols > 0:
        if n_index_cols == 1:
            index = [x[0] for x in data[:][1:]]
        else:
            transpose_index = [x[0:(n_index_cols)] for x in data[:][1:]]
            index = [[x[i] for x in transpose_index] for i in range(len(transpose_index[0]))]
        
        df = pd.DataFrame([x[n_index_cols:] for x in data[:][1:]],
                          index=index,
                          columns=data[:][0][n_index_cols:])
        
        if data[0][0] is not None:
            if n_index_cols == 1:
                df.index.name = data[0][0]
            else:
                df.index.rename([data[0][i] for i in range(n_index_cols)],inplace=True)
    else:
        df = pd.DataFrame([x for x in data[:][1:]]
                            ,columns = data[:][0][n_index_cols:])
        
        
    return df
if __name__ == "__main__":
    rng = used_range(sht)[0]
    data = xw.Range(rng).value
    df = to_df(sht,1)

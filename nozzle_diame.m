filename='C:\Users\86187\Desktop\激波雾化器\1\计算表格0917.xlsx'
sheetname='喷嘴直径计算'
row_column_in='A2:S21'
row_column_out='U2:AM21'
diameter_in=xlsread(filename,sheetname,row_column_in)
diameter_out=xlsread(filename,sheetname,row_column_out)

%range(1,1)-(20,19)
for i=1:1
    for j=1:1
        row=i
        column=j
        in = diameter_in(row,column)
        out = diameter_out(row,column)
        minus = abs(in-out)
        while minus>0.01
            in = in + 0.01
            diameter_in(row,column) = in
            xlswrite(filename,diameter_in,sheetname,row_column_in)
            diameter_out=xlsread(filename,sheetname,row_column_out)
            out = diameter_out(row,column)
            minus = out - in
        end
    end
end

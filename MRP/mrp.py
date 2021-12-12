import pandas as pd

def getdata(docname):
    Data = pd.DataFrame(pd.read_excel(docname,header=0,index_col=0))
    return Data

def data():
    X_bom = getdata("/Users/guiletong/Desktop/MRP/X_bom.xlsx")
    Y_bom = getdata("/Users/guiletong/Desktop/MRP/Y_bom.xlsx")
    TSS = getdata("/Users/guiletong/Desktop/MRP/LTandLSandStock.xlsx")
    B_IRP = getdata("/Users/guiletong/Desktop/MRP/B-IRP.xlsx")
    F_IRP = getdata("/Users/guiletong/Desktop/MRP/F-IRP.xlsx")
    X = getdata("/Users/guiletong/Desktop/MRP/X.xlsx")
    Y = getdata("/Users/guiletong/Desktop/MRP/Y.xlsx")
    B_PR = getdata("/Users/guiletong/Desktop/MRP/B_PR.xlsx")
    F_PR = getdata("/Users/guiletong/Desktop/MRP/F_PR.xlsx")
    return X_bom, Y_bom, TSS, B_IRP, F_IRP, X, Y, B_PR, F_PR

def initialize_mrp(X_bom, Y_bom, TSS, IRP, X, Y, name, PR):
    Smrp = getdata("/Users/guiletong/Desktop/MRP/StandardMRP.xlsx")
    T_x = X.columns.size
    for i in range(0,T_x): 
        if X.iloc[0,i]>0 and X.iloc[0,i]>TSS.loc["X","Stock"]: 
            Demand_X = X.iloc[0,i]-TSS.loc["X","Stock"]
            for ls in range(1,100):
                if ls*TSS.loc["X","LS"]>=Demand_X:
                    D = ls*TSS.loc["X","LS"]
                    break
            lt_X = TSS.loc["X","LT"]
            if i-lt_X<=0:
                print("False")
            Smrp.iloc[0,i-lt_X] += X_bom.loc[name,"X"]*D
            if X_bom.loc[name,"D"]>0: 
                if TSS.loc["D","Stock"]<Demand_X*X_bom.loc["D","X"]:
                    for ls in range(1,100):
                        if ls*TSS.loc["D","LS"]>=Demand_X*X_bom.loc["D","X"]:
                            D_d = ls*TSS.loc["D","LS"]
                            break
                    lt_D = TSS.loc["D","LT"]+lt_X
                    if i-lt_D<=0:
                        print("False")
                    Smrp.iloc[0,i-lt_D] += X_bom.loc[name,"D"]*D_d
    T_y = Y.columns.size
    for i in range(0,T_y): 
        if Y.iloc[0,i]>0 and Y.iloc[0,i]>TSS.loc["Y","Stock"]: 
            Demand_Y = Y.iloc[0,i]-TSS.loc["Y","Stock"]
            for ls in range(1,100):
                if ls*TSS.loc["Y","LS"]>=Demand_Y:
                    D = ls*TSS.loc["Y","LS"]
                    break
            lt_Y = TSS.loc["Y","LT"]
            if i-lt_Y<=0:
                print("False")
            Smrp.iloc[0,i-lt_Y] += Y_bom.loc[name,"Y"]*D
            if Y_bom.loc[name,"E"]>0: 
                if TSS.loc["E","Stock"]<Demand_Y*Y_bom.loc["E","Y"]:
                    for ls in range(1,100):
                        if ls*TSS.loc["E","LS"]>=Demand_Y*Y_bom.loc["E","Y"]:
                            D_e = ls*TSS.loc["E","LS"]
                            break
                    lt_E = TSS.loc["E","LT"]+lt_Y
                    if i-lt_E<=0:
                        print("False")
                    Smrp.iloc[0,i-lt_E] += Y_bom.loc[name,"E"]*D_e
    T_I = IRP.columns.size
    for i in range(0,T_I):
        Smrp.iloc[0,i] += IRP.iloc[0,i]
        Smrp.iloc[1,i] += PR.iloc[0,i]
    return Smrp

def mrp(TSS, PR, name, IRP):
    Imrp = initialize_mrp(X_bom, Y_bom, TSS, IRP, X, Y, name, PR)
    T = Imrp.columns.size
    lt = TSS.loc[name,"LT"]
    for i in range(0,T):
        if i == 0:
            if TSS.loc[name,"Stock"]+Imrp.iloc[1,i]>=Imrp.iloc[0,i]:
                Imrp.iloc[4,i] =TSS.loc[name,"Stock"]- Imrp.iloc[0,i]+Imrp.iloc[1,i]
            else:
                Imrp.iloc[2,i] = Imrp.iloc[0,i] - TSS.loc[name,"Stock"]-Imrp.iloc[1,i]
                print("需要调整计划")
        else: 
            if Imrp.iloc[4,i-1]+Imrp.iloc[1,i]>=Imrp.iloc[0,i]:
                Imrp.iloc[4,i] = Imrp.iloc[4,i-1]+Imrp.iloc[1,i] - Imrp.iloc[0,i]
            else:
                Imrp.iloc[2,i] = Imrp.iloc[0,i] - (Imrp.iloc[4,i-1]+Imrp.iloc[1,i])
                for j in range(1,100):
                    if TSS.loc[name,"LS"]*j>=Imrp.iloc[2,i]:
                        Imrp.iloc[3,i] = TSS.loc[name,"LS"]*j
                        break
                Imrp.iloc[4,i] = Imrp.iloc[3,i]- Imrp.iloc[2,i]
                Imrp.iloc[5,i-lt] = Imrp.iloc[3,i]
    return Imrp

if __name__ == '__main__':
    X_bom, Y_bom, TSS, B_IRP, F_IRP, X, Y, B_PR, F_PR = data()
    B_Mrp = mrp(TSS, B_PR, "B", B_IRP)
    F_Mrp = mrp(TSS, F_PR, "F", F_IRP)
    B_Mrp.to_excel('/Users/guiletong/Desktop/MRP/B的物料需求计划.xlsx')
    F_Mrp.to_excel('/Users/guiletong/Desktop/MRP/F的物料需求计划.xlsx')
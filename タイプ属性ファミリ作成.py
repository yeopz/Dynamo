# 必要なライブラリをインポートする
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Architecture import *
from System import Guid

# ドキュメントを取得
doc = DocumentManager.Instance.CurrentDBDocument

# ミリメートルからフィートへの変換関数
def mmToFeet(mm):
    return mm / 304.8

# 新しいファミリタイプを作成する関数
def createFamilyType(family, typeName, parameters):
    # トランザクション開始
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    # 新しいファミリタイプを作成
    newType = None
    familyTypes = family.GetFamilySymbolIds()
    for typeId in familyTypes:
        typeSymbol = doc.GetElement(typeId)
        if typeSymbol.Name == typeName:
            newType = typeSymbol
            break
    if newType is None:
        newType = family.Duplicate(typeName)
    
    # パラメータ設定
    for i in range(0, len(parameters), 2):
        paramName = parameters[i]
        paramValue = parameters[i+1]
        for param in newType.Parameters:
            if param.Definition.Name == paramName:
                # 数値パラメータの場合、単位を変換して設定
                param.Set(mmToFeet(float(paramValue)))
    
    # トランザクション終了
    TransactionManager.Instance.TransactionTaskDone()
    return newType

# Excelデータからファミリタイプを作成
excelData = IN[0]  # エクセルデータはIN[0]から入力
selectedFamily = UnwrapElement(IN[1])  # 選択したファミリはIN[1]から入力

createdTypes = []
for row in excelData:
    typeName = row[0]
    parameters = row[1:]
    createdType = createFamilyType(selectedFamily, typeName, parameters)
    createdTypes.append(createdType.Name)

# 作成したファミリタイプの名前を出力
OUT = createdTypes

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
 
# ミリメートルをフィートに変換する関数
def mmToFeet(mm):
    return mm / 304.8
 
# 屋根の境界と最高点を取得する関数
def getRoofBoundaryAndHighestPoint(roof):
    options = Options()
    geometryElement = roof.get_Geometry(options)
    minX = minY = maxX = maxY = maxZ = None
    for geometryObject in geometryElement:
        if isinstance(geometryObject, Solid):
            for face in geometryObject.Faces:
                for edgeArray in face.EdgeLoops:
                    for edge in edgeArray:
                        curve = edge.AsCurve()
                        start, end = curve.GetEndPoint(0), curve.GetEndPoint(1)
                        if minX is None or start.X < minX: minX = start.X
                        if minY is None or start.Y < minY: minY = start.Y
                        if maxX is None or start.X > maxX: maxX = start.X
                        if maxY is None or start.Y > maxY: maxY = start.Y
                        if maxZ is None or start.Z > maxZ: maxZ = start.Z
                        if end.X < minX: minX = end.X
                        if end.Y < minY: minY = end.Y
                        if end.X > maxX: maxX = end.X
                        if end.Y > maxY: maxY = end.Y
                        if end.Z > maxZ: maxZ = end.Z
    return minX, minY, maxX, maxY, maxZ
 
# ドキュメントと要素の取得
doc = DocumentManager.Instance.CurrentDBDocument
roof = UnwrapElement(IN[0])
panelInstance = UnwrapElement(IN[1])
 
if not isinstance(panelInstance, FamilyInstance):
    raise Exception("The solar panel input must be of FamilyInstance type.")
 
panelType = panelInstance.Symbol
 
# パネルの寸法と間隔をミリメートルで指定
panelWidthMM, panelHeightMM = 1096, 1754
x_spacingMM, y_spacingMM = 0, 0
panelOffsetMM, offsetMM, wiring_spaceMM = 254, 1500, 1000
 
# フィートに変換
panelWidth, panelHeight = mmToFeet(panelWidthMM), mmToFeet(panelHeightMM)
x_spacing, y_spacing = mmToFeet(x_spacingMM), mmToFeet(y_spacingMM)
panelOffset, offset, wiring_space = mmToFeet(panelOffsetMM), mmToFeet(offsetMM), mmToFeet(wiring_spaceMM)
 
minX, minY, maxX, maxY, maxZ = getRoofBoundaryAndHighestPoint(roof)
 
# オフセットを適用
minX += offset
minY += offset
maxX -= offset
maxY -= offset
 
TransactionManager.Instance.EnsureInTransaction(doc)
placedPanels = []
 
# 屋根の利用可能な長さと幅
usableLength = maxX - minX
usableWidth = maxY - minY
 
# X軸とY軸に配置可能なパネルの数を計算
panelsX = int(usableLength / (panelWidth + x_spacing))
panelsY = int(usableWidth / (panelHeight + y_spacing))
 
# 配線スペースの数を計算（IN[3]とIN[4]から取得）
wiring_spacesX = int(IN[3])
wiring_spacesY = int(IN[4])
 
# 実際に使用するスペーシングを計算
actual_spacingX = (usableLength - panelsX * panelWidth - wiring_spacesX * wiring_space) / max(panelsX - 1, 1)
actual_spacingY = (usableWidth - panelsY * panelHeight - wiring_spacesY * wiring_space) / max(panelsY - 1, 1)
 
# パネルの配置
x = minX
while x + panelWidth <= maxX:
    y = minY
    while y + panelHeight <= maxY:
        # パネルの配置
        newLocation = XYZ(x + panelWidth / 2, y + panelHeight / 2, maxZ + panelOffset)
        newPanel = doc.Create.NewFamilyInstance(newLocation, panelType, roof, Structure.StructuralType.NonStructural)
        placedPanels.append(newPanel)
        y += panelHeight + actual_spacingY
    x += panelWidth + actual_spacingX
 
TransactionManager.Instance.TransactionTaskDone()
 
# 総価格の計算
panel_count = len(placedPanels)
panel_unit_price = int(IN[2])
total_price = panel_count * panel_unit_price
 
# 出力
OUT = {'配置されたパネルの総数': panel_count, 'ソーラーパネルの総額': total_price}
import adsk.core
import adsk.fusion
import adsk.cam
import traceback
import json
import re
import os
import os.path
import datetime
import sys
from .Modules import xlsxwriter



# Copy the folder to: %AppData%\Autodesk\Autodesk Fusion 360\API\AddIns
# (This is usually: C:\Users\userName\AppData\Autodesk\Autodesk Fusion 360\API\AddIns

# Global list to keep all event handlers in scope.
# This is only needed with Python.
handlers = []
app = adsk.core.Application.get()
ui = app.userInterface
cmdId = "BomAddInMenu"
cmdName = "Bill Of Materials"
dialogTitle = "Create BOM->Excel"
cmdDesc = "Tworzy listę elementów (BOM) oraz zapisuje ją do pliku XLSX lub CSV."
cmdRes = ".//resources//BOM-Excel"
panelID = "SolidCreatePanel"
name_logo_file = "logoBOM"

# Event handler for the commandCreated event.
class SampleCommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    global cmdId
    global ui

    def __init__(self):
        super().__init__()

    def openFileLogo(self, file):
        try:
            strlogo_name = open(file, 'r').read()
            return  strlogo_name
        except:  
            return "Fusion 360"

    def notify(self, args):
        product = app.activeProduct
        design = adsk.fusion.Design.cast(product)
        lastPrefs = design.attributes.itemByName(cmdId, "lastUsedOptions")
        _onlySelectedComps = False
        _ignoreCompsWithoutBodies = True
        _ignoreLinkedComps = False
        _ignoreVisibleState = True
        _ignoreUnderscorePrefixedComps = True
        _underscorePrefixStrip = False
        _fullList = True
        _sortDims = False
        _openFile = True
        _dataCSV = True
        _nameProj = True
        _includeArea = False
        _includeMass = False
        _includeDensity = False
        _includeMaterial = False
        _includeDesc = True
        _fileType = True
        _stringlogo = self.openFileLogo(name_logo_file)
        _decimalPlaces = True

        if lastPrefs:
            lastPrefs = json.loads(lastPrefs.value)
            _onlySelectedComps = lastPrefs.get("onlySelComp", False)
            _ignoreCompsWithoutBodies = lastPrefs.get("ignoreCompWoBodies", True)
            _ignoreLinkedComps = lastPrefs.get("ignoreLinkedComp", True)
            _ignoreVisibleState = lastPrefs.get("ignoreVisibleState", True)
            _ignoreUnderscorePrefixedComps = lastPrefs.get("ignoreUnderscorePrefComp", True)
            _underscorePrefixStrip = lastPrefs.get("underscorePrefixStrip", False)            
            _sortDims = lastPrefs.get("sortDims", False)
            _openFile = lastPrefs.get("openFile", True)
            _dataCSV = lastPrefs.get("dataCSV", False)
            _nameProj = lastPrefs.get("nameProj", False)
            _fullList = lastPrefs.get("fullList", False)
            _includeDesc = lastPrefs.get("includeDesc", False)
            _includeArea = lastPrefs.get("includeArea", False)
            _includeMass = lastPrefs.get("includeMass", False)
            _includeDensity = lastPrefs.get("includeDensity", False)
            _includeMaterial = lastPrefs.get("includeMaterial", False)
            _fileType = lastPrefs.get("fileType", False)
            _stringlogo = lastPrefs.get("stringlogo", False)
            _decimalPlaces = lastPrefs.get("decimalPlaces", True)

        try:    
            eventArgs = adsk.core.CommandCreatedEventArgs.cast(args)
            # Get the command
            cmd = eventArgs.command
            # Get the CommandInputs collection to create new command inputs.            
            inputs = cmd.commandInputs

            ipSelectComps = inputs.addBoolValueInput(cmdId + "_onlySelectedComps", "Tylko wybrane", True, "", _onlySelectedComps)
            ipSelectComps.tooltip = "Zostaną użyte tylko wybrane komponenty"

            ipWoBodies = inputs.addBoolValueInput(cmdId + "_ignoreCompsWithoutBodies", "Wyklucz, jeśli nie ma ciał", True, "", _ignoreCompsWithoutBodies)
            ipWoBodies.tooltip = "Wyklucz wszystkie komponenty, jeśli mają co najmniej jedno ciało"

            ipLinkedComps = inputs.addBoolValueInput(cmdId + "_ignoreLinkedComps", "Wyklucz dołączone", True, "", _ignoreLinkedComps)
            ipLinkedComps.tooltip = "Wyklucz wszystkie komponenty, które są połączone z projektem"

            ipVisibleState = inputs.addBoolValueInput(cmdId + "_ignoreVisibleState", "Ignoruje stan widoczności", True, "", _ignoreVisibleState)
            ipVisibleState.tooltip = "Ignoruje widoczność elementu"

            ipUnderscorePrefix = inputs.addBoolValueInput(cmdId + "_ignoreUnderscorePrefixedComps", 'Wyklucz "_"', True, "", _ignoreUnderscorePrefixedComps)
            ipUnderscorePrefix.tooltip = 'Wyklucz wszystkie komponenty, których nazwa zaczyna się od "_"'

            ipUnderscorePrefixStrip = inputs.addBoolValueInput(cmdId + "_underscorePrefixStrip", 'Usuń "_"', True, "", _underscorePrefixStrip)
            ipUnderscorePrefixStrip.tooltip = 'Jeśli zaznaczone, "_" jest usuwane z nazwy komponentów'
            ipUnderscorePrefixStrip.isVisible = not _ignoreUnderscorePrefixedComps            

            infullList = inputs.addBoolValueInput(cmdId + '_fullList', 'Zwarta lista', True, '', _fullList)
            infullList.tooltip = "Kasuje puste wpisy oraz jeśli elementy powtarzające sie mają takie same wymiary\n to je sumuje i pokazuje jako jedna pozycja. "
            #infullList.isVisible = True
            
            ipsortDims = inputs.addBoolValueInput(cmdId + '_sortDims', 'Sortowanie wymiarów', True, '', _sortDims)
            ipsortDims.tooltip = "Sortuje wymiary tak aby najdłuższy wymiar był jako długość."
            #ipsortDims.isVisible = True

            ipdecimalPlaces = inputs.addIntegerSpinnerCommandInput(cmdId + '_decimalPlaces', 'Miejsca po przecinku', 0, 5, 1, 0)
            ipdecimalPlaces.tooltip = "Dane w BOM będą podawane z taką ilością miejsc po przecinku."

            grpFile = inputs.addGroupCommandInput(cmdId + '_grpFile', 'PLIK')
            grpFileChildren = grpFile.children
            
            inOpenFile = grpFileChildren.addBoolValueInput(cmdId + '_openFile', 'Otwórz plik', True, '', _openFile)
            inOpenFile.tooltip = "Otwiera automatycznie plik po utworzeniu."

            inFileType = grpFileChildren.addDropDownCommandInput(cmdId + '_fileType', 'Rodzaj pliku', adsk.core.DropDownStyles.TextListDropDownStyle);
            fileItems = inFileType.listItems
            fileItems.add('Excel', True, '')
            fileItems.add('CSV', False, '')

            grpPhysics = inputs.addGroupCommandInput(cmdId + '_grpPhysics', 'DOŁĄCZ')
            # if _dataCSV or _nameProj or _includeDesc or _includeArea or _includeMass or _includeDensity or _includeMaterial:
            #     grpPhysics.isExpanded = True
            # else:
            # 	grpPhysics.isExpanded = False
            grpPhysicsChildren = grpPhysics.children

            strInput = grpPhysicsChildren.addStringValueInput(cmdId + '_stringlogo', 'Logo', _stringlogo)

            inDataCSV = grpPhysicsChildren.addBoolValueInput(cmdId + '_dataCSV', 'Data utworzenia BOM', True, '', _dataCSV)
            inDataCSV.tooltip = "Dołącza datę utworzenia BOM do pliku."
            #inDataCSV.isVisible = True

            inNameProj = grpPhysicsChildren.addBoolValueInput(cmdId + '_nameProj', 'Nazwa projektu', True, '', _nameProj)
            inNameProj.tooltip = "Dołącza nazwę projektu do pliku."
            #inNameProj.isVisible = True

            ipIncludeArea = grpPhysicsChildren.addBoolValueInput(cmdId + "_includeArea", "Powierzchnia", True, "", _includeArea)
            ipIncludeArea.tooltip = "Dołącza powierzchnię komponentu w cm^2"

            ipIncludeMass = grpPhysicsChildren.addBoolValueInput(cmdId + "_includeMass", "Waga", True, "", _includeMass)
            ipIncludeMass.tooltip = "Dołącza masę komponentu w kg"

            ipIncludeDensity = grpPhysicsChildren.addBoolValueInput(cmdId + "_includeDensity", "Gęstość", True, "", _includeDensity)
            ipIncludeDensity.tooltip = "Dołącza gęstość komponentu w kg/cm^3"

            ipIncludeMaterial = grpPhysicsChildren.addBoolValueInput(cmdId + "_includeMaterial", "Materiał", True, "", _includeMaterial)
            ipIncludeMaterial.tooltip = "Dołącza materiał"

            ipCompDesc = grpPhysicsChildren.addBoolValueInput(cmdId + '_includeDesc', 'Opis', True, '', _includeDesc)
            ipCompDesc.tooltip = "Zawiera opis komponentu. Możesz dodać opis, klikając komponent\n prawym przyciskiem myszy i otwierając panel Właściwości."
            
            
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

        # Connect to the execute event.
        onExecute = SampleCommandExecuteHandler()
        cmd.execute.add(onExecute)
        handlers.append(onExecute)

        # Connect to the inputChanged event.
        onInputChanged = SampleCommandInputChangedHandler()
        cmd.inputChanged.add(onInputChanged)
        handlers.append(onInputChanged)


# Event handler for the inputChanged event.
class SampleCommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()
    def notify(self, args):
        global ui
        global cmdId

        # eventArgs = adsk.core.InputChangedEventArgs.cast(args)
        command = args.firingEvent.sender
        inputs = command.commandInputs

        if inputs.itemById(cmdId + "_ignoreUnderscorePrefixedComps").value is True:
            inputs.itemById(cmdId + "_underscorePrefixStrip").isVisible = False
        else:
            inputs.itemById(cmdId + "_underscorePrefixStrip").isVisible = True


# Event handler for the execute event.
class SampleCommandExecuteHandler(adsk.core.CommandEventHandler):
    global cmdId
    def __init__(self):
        super().__init__()
        
    def replacePointDelimterOnPref(self, pref, value):
            if (pref):
                return str(value).replace(".", ",")
            return str(value)

    def getDataTime(self):
        now = datetime.date.today()
        return now.strftime("%d-%m-%Y")

    def openFile(self, file):
        try:
            strlogo_name = open(file, 'r').read()
            return  strlogo_name
        except OSError as e:  ## if failed, report it back to the user ##
            return (e.filename +'\n ' + e.strerror)

    def formatDecimal(self, file, decimal):
        if decimal == 0:
            return "{0:.0f}".format(file)
        elif decimal == 1:
            return "{0:.1f}".format(file)
        elif decimal == 2:
            return "{0:.2f}".format(file)
        elif decimal == 3:
            return "{0:.3f}".format(file)
        elif decimal >= 4:
            return "{0:.4f}".format(file)

    def collectDataExcel(self, design, bom, prefs, filename):
        defaultUnit = design.fusionUnitsManager.defaultLengthUnits        
        # Document name
        app = adsk.core.Application.get()
        docNameWithVersion = app.activeDocument.name
        docName = docNameWithVersion.rsplit(' ',1)[0]

        self.saveFile(name_logo_file, prefs["stringlogo"])     
        logo_name = self.openFile(name_logo_file)

        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet(logo_name)
        worksheet.set_tab_color('red')
        # worksheet2.set_tab_color('green')
        # worksheet3.set_tab_color('#FF9900')  # Orange

        # Widen the first column to make the text clearer.
        worksheet.set_column('A:A', 20)
        # Set up some formats to use.

        max_row = len(bom)
        #######################################################################
        #
        # Example 1. Freeze pane on the top row.
        #
        
        worksheet.freeze_panes(4, 1)
        worksheet.autofilter(3, 0, max_row, 4)
        #######################################################################
        #
        # Set up some formatting and text to highlight the panes.
        #
        bold = workbook.add_format({'bold': True})
        italic = workbook.add_format({'italic': True})
        align_left = workbook.add_format({'align': 'left'})
        align_center = workbook.add_format({'align': 'center'})
        bold_center = workbook.add_format({'bold': True,
                                            'align': 'center'})
        name_format = workbook.add_format({'bold': True,
                                            'align': 'left',
                                            'valign': 'vcenter',
                                            'fg_color': '#d2f2c7'})        
        instances_format = workbook.add_format({'bold': True,
                                            'align': 'center',
                                            'valign': 'vcenter',
                                            'fg_color': '#ded38e'})        
        header_format = workbook.add_format({'bold': True,
                                            'align': 'center',
                                            'valign': 'vcenter',
                                            'fg_color': '#D7E4BC',
                                            'border': 1})        
        merge_format = workbook.add_format({'bold': 1,
                                            'border': 1,
                                            'align': 'center',
                                            'font_size': 30,
                                            'valign': 'vcenter',
                                            'fg_color': '#D7E4BC'})

        # Other sheet formatting.
        worksheet.set_column('A:A', 50)
        worksheet.set_column('B:B', 10)
        worksheet.set_column('C:E', 18)
        worksheet.set_row(0, 20)
        worksheet.set_selection('A4')


        row = 0
        column = 1

        worksheet.merge_range('A1:A3', logo_name, merge_format)

        if prefs["dataCSV"]:
            worksheet.write_rich_string(row, column,  "Data utworzenia dokumentu: ", bold, self.getDataTime())
        if prefs["nameProj"]:
            row +=1
            if prefs["onlySelComp"]:
                worksheet.write(row, column, "UWAGA!",bold)
                worksheet.write_rich_string(row + 1, column, "Utworzono na podstawie wybranych elementów z projektu: ", bold, docName)
            else:
                worksheet.write_rich_string(row, column, "Utworzono na podstawie projektu: ", bold, docName)
        column = 0        
        row = 3
        worksheet.write(row, column, "NAZWA ELEMENTU", header_format)        
        column += 1
        worksheet.write(row, column, "ILOŚĆ", header_format)
        column += 1
        worksheet.write(row, column, "SZEROKOŚĆ [" + defaultUnit + "]", header_format)
        column += 1
        worksheet.write(row, column, "DŁUGOŚĆ [" + defaultUnit + "]", header_format)
        column += 1
        worksheet.write(row, column, "WYSOKOŚĆ [" + defaultUnit + "]", header_format)
        column += 1
        if prefs["includeArea"]:
            worksheet.set_column(column, column, 20)
            worksheet.write(row, column, "POWIERZCHNIA [cm^2]", header_format)
            column += 1
        if prefs["includeMass"]:
            worksheet.set_column(column, column, 20)
            worksheet.write(row, column, "CIĘŻAR [kg]", header_format)
            column += 1
        if prefs["includeDensity"]:
            worksheet.set_column(column, column, 20)
            worksheet.write(row, column, "GĘSTOŚĆ [kg/cm^2]", header_format)
            column += 1
        if prefs["includeMaterial"]:
            worksheet.set_column(column, column, 30)
            worksheet.write(row, column, "MATERIAŁ", header_format)
            column += 1
        if prefs["includeDesc"]:
            worksheet.set_column(column, column, 70)
            worksheet.write(row, column, "OPIS", header_format)
            column += 1

        double_bom =  []
        
        for item in bom:
            dimX = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["x"], defaultUnit, False))
            dimY = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["y"], defaultUnit, False))
            dimZ = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["z"], defaultUnit, False))
            
            dim = 0
            for k in item["boundingBox"]:
                dim += item["boundingBox"][k]
            if dim > 0:                
                dimSorted = sorted([dimX, dimY, dimZ])
                if prefs["sortDims"]:
                    dimSorted = sorted([dimX, dimY, dimZ])
                    bbZ = self.formatDecimal(dimSorted[2], prefs["decimalPlaces"])
                    bbX = self.formatDecimal(dimSorted[0], prefs["decimalPlaces"])
                    bbY = self.formatDecimal(dimSorted[1], prefs["decimalPlaces"])
                else:
                    bbX = self.formatDecimal(dimX, prefs["decimalPlaces"])
                    bbY = self.formatDecimal(dimY, prefs["decimalPlaces"])
                    bbZ = self.formatDecimal(dimZ, prefs["decimalPlaces"])                    
           
            name = self.filterFusionCompNameInserts(item["name"])
            append = True
            for it in double_bom: 
                if prefs["fullList"]:
                    if name == it["double_name"] and bbX == it["double_dimX"] and bbY == it["double_dimY"] and bbZ == it["double_dimZ"]:
                        it["double_instances"] = it["double_instances"] + item["instances"]
                        append = False
            if append:                           
                double_bom.append({
                        # "double_component": comp,
                        "double_name": name,
                        "double_instances": item["instances"],
                        # " double_volume": self.getBodiesVolume(comp.bRepBodies),
                        "double_dimX": bbX,
                        "double_dimY": bbY,
                        "double_dimZ": bbZ,
                        "double_area": item["area"],
                        "double_mass": item["mass"],
                        "double_density": item["density"],
                        "double_material": item["material"],
                        "double_desc": item["desc"]
                    })

        column = 0
        for  double_item in double_bom:
            row += 1
            name = self.filterFusionCompNameInserts(double_item["double_name"])
            worksheet.write(row, column, name, name_format)
            column += 1
            worksheet.write(row, column, double_item["double_instances"], instances_format)            
            column += 1
            worksheet.write(row, column, double_item["double_dimX"], align_center)            
            column += 1
            worksheet.write(row, column, double_item["double_dimY"], align_center)            
            column += 1
            worksheet.write(row, column, double_item["double_dimZ"], align_center)            
            if prefs["includeArea"]:
                column += 1                
                worksheet.write(row, column, self.formatDecimal(double_item["double_area"], prefs["decimalPlaces"]), align_left)  
            if prefs["includeMass"]:
                column += 1
                worksheet.write(row, column, self.formatDecimal(double_item["double_mass"], prefs["decimalPlaces"]), align_left)  
            if prefs["includeDensity"]:
                column += 1
                worksheet.write(row, column, self.formatDecimal(double_item["double_density"], prefs["decimalPlaces"]), align_left)  
            if prefs["includeMaterial"]:
                column += 1
                worksheet.write(row, column, double_item["double_material"], align_left)  
            if prefs["includeDesc"]:
                column += 1
                worksheet.write(row, column, double_item["double_desc"], align_left)            
            column = 0
        workbook.close() 


    def collectData(self, design, bom, prefs):
        defaultUnit = design.fusionUnitsManager.defaultLengthUnits        
        # Document name
        app = adsk.core.Application.get()
        docNameWithVersion = app.activeDocument.name
        docName = docNameWithVersion.rsplit(' ',1)[0]

        csvStr = ''
        if prefs["dataCSV"]:
            csvStr += '"' + "Data utworzenia dokumentu: " + self.getDataTime() + '",\n\n'
        if prefs["nameProj"]:
            csvStr += '"' + "Utworzono na podstawie projektu: " + docName + '",\n\n\n'

        csvHeader = ["NAZWA ELEMENTU", "ILOŚĆ"]

        if prefs["sortDims"]:
            csvHeader.append("WYSOKOŚĆ [" + defaultUnit + "]")
            csvHeader.append("SZEROKOŚĆ [" + defaultUnit + "]")
            csvHeader.append("DŁUGOŚĆ [" + defaultUnit + "]") 
        else:
            csvHeader.append("SZEROKOŚĆ [" + defaultUnit + "]")
            csvHeader.append("DŁUGOŚĆ [" + defaultUnit + "]")
            csvHeader.append("WYSOKOŚĆ [" + defaultUnit + "]")
        if prefs["includeArea"]:
            csvHeader.append("POWIERZCHNIA [cm^2]")
        if prefs["includeMass"]:
            csvHeader.append("CIĘŻAR [kg]")
        if prefs["includeDensity"]:
            csvHeader.append("GĘSTOŚĆ [kg/cm^2]")
        if prefs["includeMaterial"]:
            csvHeader.append("MATERIAŁ")        
        if prefs["includeDesc"]:
            csvHeader.append("OPIS") 
     
        for k in csvHeader:
            csvStr += '"' + k + '",'
        csvStr += '\n'
  
        double_bom =  []
        
        for item in bom:
            dimX = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["x"], defaultUnit, False))
            dimY = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["y"], defaultUnit, False))
            dimZ = float(design.fusionUnitsManager.formatInternalValue(item["boundingBox"]["z"], defaultUnit, False))
            
            dim = 0
            for k in item["boundingBox"]:
                dim += item["boundingBox"][k]
            if dim > 0:
                if prefs["sortDims"]:
                    dimSorted = sorted([dimX, dimY, dimZ])
                    bbZ = self.formatDecimal(dimSorted[2], prefs["decimalPlaces"])
                    bbX = self.formatDecimal(dimSorted[0], prefs["decimalPlaces"])
                    bbY = self.formatDecimal(dimSorted[1], prefs["decimalPlaces"])
                else:
                    bbX = self.formatDecimal(dimX, prefs["decimalPlaces"])
                    bbY = self.formatDecimal(dimX, prefs["decimalPlaces"])
                    bbZ = self.formatDecimal(dimX, prefs["decimalPlaces"])
            
            name = self.filterFusionCompNameInserts(item["name"])
            append = True
            for it in double_bom:
                if prefs["fullList"]:
                    if name == it["double_name"] and bbX == it["double_dimX"] and bbY == it["double_dimY"] and bbZ == it["double_dimZ"]:
                        it["double_instances"] = it["double_instances"] + item["instances"]
                        append = False              
            if append:                           
                double_bom.append({
                        "double_name": name,
                        "double_instances": item["instances"],
                        "double_dimX": bbX,
                        "double_dimY": bbY,
                        "double_dimZ": bbZ,
                        "double_area": item["area"],
                        "double_mass": item["mass"],
                        "double_density": item["density"],
                        "double_material": item["material"],
                        "double_desc": item["desc"]
                    })

        interspace = '","'
        interspace_start = '"'
        interspace_end = '",'
        
        for  double_item in double_bom:
            csvStr += interspace_start + double_item["double_name"] + interspace + self.replacePointDelimterOnPref(prefs["useComma"], double_item["double_instances"])
            csvStr += interspace + double_item["double_dimX"] + interspace + double_item["double_dimY"] + interspace + double_item["double_dimZ"]
            if prefs["includeArea"]:
                csvStr += interspace + self.replacePointDelimterOnPref(prefs["useComma"], self.formatDecimal(double_item["double_area"], prefs["decimalPlaces"]))
            if prefs["includeMass"]:
                csvStr += interspace + self.replacePointDelimterOnPref(prefs["useComma"], self.formatDecimal(double_item["double_mass"], prefs["decimalPlaces"]))
            if prefs["includeDensity"]:
                csvStr += interspace + self.replacePointDelimterOnPref(prefs["useComma"], self.formatDecimal(double_item["double_density"], prefs["decimalPlaces"]))
            if prefs["includeMaterial"]:
                csvStr += interspace + double_item["double_material"]
            if prefs["includeDesc"]:
                csvStr += interspace + double_item["double_desc"]
            csvStr += interspace_end
            csvStr += '\n'

        return csvStr

    def getPrefsObject(self, inputs):
    
            obj = {
                "onlySelComp": inputs.itemById(cmdId + "_onlySelectedComps").value,
                "ignoreLinkedComp": inputs.itemById(cmdId + "_ignoreLinkedComps").value,
                "ignoreCompWoBodies": inputs.itemById(cmdId + "_ignoreCompsWithoutBodies").value,
                "ignoreVisibleState": inputs.itemById(cmdId + "_ignoreVisibleState").value,
                "ignoreUnderscorePrefComp": inputs.itemById(cmdId + "_ignoreUnderscorePrefixedComps").value,
                "underscorePrefixStrip": inputs.itemById(cmdId + "_underscorePrefixStrip").value,
                "sortDims": inputs.itemById(cmdId + "_sortDims").value,
                "openFile": inputs.itemById(cmdId + "_openFile").value,
                "dataCSV": inputs.itemById(cmdId + "_dataCSV").value,
                "nameProj": inputs.itemById(cmdId + "_nameProj").value,                
                "fullList": inputs.itemById(cmdId + "_fullList").value,  
                "includeDesc": inputs.itemById(cmdId + "_includeDesc").value,
                "includeArea" : inputs.itemById(cmdId + "_includeArea").value,
                "includeMass" : inputs.itemById(cmdId + "_includeMass").value,
                "includeDensity" : inputs.itemById(cmdId + "_includeDensity").value,
                "includeMaterial" : inputs.itemById(cmdId + "_includeMaterial").value,
                "fileType": inputs.itemById(cmdId + "_fileType").selectedItem.name,
                "stringlogo": inputs.itemById(cmdId + "_stringlogo").value,
                "decimalPlaces": inputs.itemById(cmdId + "_decimalPlaces").value,                
                "generateCutlList": True,
                "useComma": True
            }
            return obj  

    def getBodiesVolume(self, bodies):
            volume = 0
            for bodyK in bodies:
                if bodyK.isSolid:
                    volume += bodyK.volume
            return volume

    # Calculates a tight bounding box around the input body.  An optional
    # tolerance argument is available.  This specificies the tolerance in
    # centimeters.  If not provided the best existing display mesh is used.
    def calculateTightBoundingBox(self, body, tolerance = 0):
        try:
            # If the tolerance is zero, use the best display mesh available.
            if tolerance <= 0:
                # Get the best display mesh available.
                triMesh = body.meshManager.displayMeshes.bestMesh
            else:
                # Calculate a new mesh based on the input tolerance.
                meshMgr = adsk.fusion.MeshManager.cast(body.meshManager)
                meshCalc = meshMgr.createMeshCalculator()
                meshCalc.surfaceTolerance = tolerance
                triMesh = meshCalc.calculate()
       
            # Calculate the range of the mesh.
            smallPnt = adsk.core.Point3D.cast(triMesh.nodeCoordinates[0])
            largePnt = adsk.core.Point3D.cast(triMesh.nodeCoordinates[0])
            vertex = adsk.core.Point3D.cast(None)
            for vertex in triMesh.nodeCoordinates:
                if vertex.x < smallPnt.x:
                    smallPnt.x = vertex.x
                   
                if vertex.y < smallPnt.y:
                    smallPnt.y = vertex.y
                   
                if vertex.z < smallPnt.z:
                    smallPnt.z = vertex.z
               
                if vertex.x > largePnt.x:
                    largePnt.x = vertex.x
                   
                if vertex.y > largePnt.y:
                    largePnt.y = vertex.y
                   
                if vertex.z > largePnt.z:
                    largePnt.z = vertex.z
                   
            # Create and return a BoundingBox3D as the result.
            return(adsk.core.BoundingBox3D.create(smallPnt, largePnt))
        except:
            # An error occurred so return None.
            return(None)
            
    def getBodiesBoundingBox(self, bodies):
        minPointX = maxPointX = minPointY = maxPointY = minPointZ = maxPointZ = 0
        # Examining the maximum min point distance and the maximum max point distance.       
        for body in bodies:
            if body.isSolid:
                bb = self.calculateTightBoundingBox(body, 500)
                if not bb:
                    return None
                if not minPointX or bb.minPoint.x < minPointX:
                    minPointX = bb.minPoint.x
                if not maxPointX or bb.maxPoint.x > maxPointX:
                    maxPointX = bb.maxPoint.x
                if not minPointY or bb.minPoint.y < minPointY:
                    minPointY = bb.minPoint.y
                if not maxPointY or bb.maxPoint.y > maxPointY:
                    maxPointY = bb.maxPoint.y
                if not minPointZ or bb.minPoint.z < minPointZ:
                    minPointZ = bb.minPoint.z
                if not maxPointZ or bb.maxPoint.z > maxPointZ:
                    maxPointZ = bb.maxPoint.z
                
        return {
            "x": maxPointX - minPointX,
            "y": maxPointY - minPointY,
            "z": maxPointZ - minPointZ
        } 

    def getPhysicsArea(self, bodies):
        area = 0
        for body in bodies:
            if body.isSolid:
                if body.physicalProperties:
                    area += body.physicalProperties.area
        return area

    def getPhysicalMass(self, bodies):
        mass = 0
        for body in bodies:
            if body.isSolid:
                if body.physicalProperties:
                    mass += body.physicalProperties.mass
        return mass

    def getPhysicalDensity(self, bodies):
        density = 0
        if bodies.count > 0:
            body = bodies.item(0)
            if body.isSolid:
                if body.physicalProperties:
                    density = body.physicalProperties.density
            return density

    def getPhysicalMaterial(self, bodies):
        matList = []
        for body in bodies:
            if body.isSolid and body.material:
                mat = body.material.name
                if mat not in matList:
                    matList.append(mat)
        return ', '.join(matList)
        
    def filterFusionCompNameInserts(self, name):
            name = re.sub("\([0-9]+\)$", '', name)
            name = name.strip()
            name = re.sub("v[0-9]+$", '', name)
            return name.strip()
 
    def saveFile(self, fnm, fbom):
        try:
                output = open(fnm, 'w')
                output.writelines(fbom)
                output.close()
                return 0
        except OSError as e:  ## if failed, report it back to the user ##
            return (e.filename +'\n ' + e.strerror)

    
    def notify(self, args):
        global dialogTitle
        ui = None
        try:
            app = adsk.core.Application.get()
            ui  = app.userInterface

            product = app.activeProduct
            design = adsk.fusion.Design.cast(product)
            eventArgs = adsk.core.CommandEventArgs.cast(args)            
            inputs = eventArgs.command.commandInputs
            
            if not design:
                ui.messageBox('Brak aktywnego projektu', dialogTitle)
                return
            prefs = self.getPrefsObject(inputs)
                
                
            # Get all occurrences in the root component of the active design
            root = design.rootComponent

            occs = []
            if prefs["onlySelComp"]:
                if ui.activeSelections.count > 0:
                    selections = ui.activeSelections
                    for selection in selections:
                        if (hasattr(selection.entity, "objectType") and selection.entity.objectType == adsk.fusion.Occurrence.classType()):
                            occs.append(selection.entity)
                            if selection.entity.component:
                                for item in selection.entity.component.allOccurrences:
                                    occs.append(item)
                        else:
                            ui.messageBox('No components selected!\nPlease select some components.')
                            return
                else:
                    ui.messageBox('No components selected!\nPlease select some components.')
                    return
            else:
                occs = root.allOccurrences


            
            if len(occs) == 0:
                ui.messageBox('W tym projekcie nie ma żadnych komponentów.')
                return
            
            # Set styles of progress dialog.
            progressDialog = ui.createProgressDialog()
            progressDialog.cancelButtonText = 'Zakończ'
            progressDialog.isBackgroundTranslucent = False
            progressDialog.isCancelButtonShown = True
            
            steps = 0
            for occ1 in occs:
                occ1.component
                steps += 1
            
            # Show dialog
            progressDialog.show('Postęp', 'Procent: %p%, Aktualny element: %v', 0, steps)
                    
            # Gather information about each unique component
            bom = []
            
            for occ in occs:
                if progressDialog.wasCancelled:
                        break
                progressDialog.progressValue += 1
                comp = occ.component
                if comp.name.startswith('_') and prefs["ignoreUnderscorePrefComp"]:
                    continue
                elif prefs["ignoreLinkedComp"] and design != comp.parentDesign:
                    continue
                elif not comp.bRepBodies.count and prefs["ignoreCompWoBodies"]:
                    continue
                elif not occ.isVisible and prefs["ignoreVisibleState"] is False:
                    continue
                else:
                    jj = 0
                    for bomI in bom:
                        if bomI['component'] == comp:
                            # Increment the instance count of the existing row.
                            bomI['instances'] += 1
                            break
                        jj += 1                    

                    if jj == len(bom):
                        # Add this component to the BOM
                        bb = self.getBodiesBoundingBox(comp.bRepBodies)
                        if not bb:
                            if ui:
                                ui.messageBox('Nie wszystkie moduły Fusion są jeszcze załadowane, kliknij element główny, aby je załadować i spróbuj ponownie.')
                            return
                            
                        bom.append({
                            "component": comp,
                            "name": comp.name,
                            "instances": 1,
                            "volume": self.getBodiesVolume(comp.bRepBodies),
                            "boundingBox": bb,
                            "area": self.getPhysicsArea(comp.bRepBodies),
                            "mass": self.getPhysicalMass(comp.bRepBodies),
                            "density": self.getPhysicalDensity(comp.bRepBodies),
                            "material": self.getPhysicalMaterial(comp.bRepBodies),
                            "desc": comp.description
                        })
                    
            # Hide the progress dialog at the end.
            progressDialog.hide()

            fileSaveType = str(prefs["fileType"])
            if fileSaveType == 'Excel':
                saveExcel = True
            else:
                saveExcel = False
            
            fileDialog = ui.createFileDialog()
            fileDialog.isMultiSelectEnabled = False
            fileDialog.title = dialogTitle + " filename"
            if saveExcel:
                fileDialog.filter = 'XLSX (*.xlsx)'
            else:
                fileDialog.filter = 'CSV (*.csv)'                
            fileDialog.filterIndex = 0
            dialogResult = fileDialog.showSave()
            if dialogResult == adsk.core.DialogResults.DialogOK:
                filename = fileDialog.filename
            else:
                return


            if saveExcel:

                self.collectDataExcel(design, bom, prefs, filename)

            else:
                bomStr = self.collectData(design, bom, prefs)                            
                checkFilename = self.saveFile(filename, bomStr)
            
                if checkFilename == 0:                
                    # Save last chosen options    
                    design.attributes.add(cmdId, "lastUsedOptions", json.dumps(prefs))                
                else:   
                    message = "BŁĄD ZAPISU PLIKU!!!\n \"" + str(checkFilename) + "\" \n"
                    ui.messageBox(message)
                    return
           
            if prefs["openFile"]:
                os.startfile(filename)
            else:
                ui.messageBox('Zapisano do pliku "' + filename + '"')
                
            design.attributes.add(cmdId, "lastUsedOptions", json.dumps(prefs)) 
            

        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))

def run(context):
    global ui
    global cmdId
    global dialogTitle
    global cmdDesc
    global cmdRes
    global panelID

    try:  
        # Get the CommandDefinitions collection.
        cmdDefs = ui.commandDefinitions
        
        # Create a button command definition.
        buttonSample = cmdDefs.addButtonDefinition(cmdId, dialogTitle, cmdDesc, cmdRes)
        
        # Connect to the command created event.
        sampleCommandCreated = SampleCommandCreatedEventHandler()
        buttonSample.commandCreated.add(sampleCommandCreated)
        handlers.append(sampleCommandCreated)
        
        # Get the ADD-INS panel in the model workspace. 
        addInsPanel = ui.allToolbarPanels.itemById(panelID)
        
        # Add the button to the bottom of the panel.
        buttonControl = addInsPanel.controls.addCommand(buttonSample, "")
        buttonControl.isVisible = True
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    try:
        global app
        global ui
        global cmdId
        global panelID
        
        # Clean up the UI.
        cmdDef = ui.commandDefinitions.itemById(cmdId)
        if cmdDef:
            cmdDef.deleteMe()
            
        addinsPanel = ui.allToolbarPanels.itemById(panelID)
        cntrl = addinsPanel.controls.itemById(cmdId)
        if cntrl:
            cntrl.deleteMe()
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
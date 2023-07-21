"""
openLCA Version: 2.0.0

Description: 

	This script will run the impact calculation for the 'LCO2 life cycle stages' model and insert the results into a LCO2 Model excel template file which you can selected when starting the script. The script uses the impact method 'EF 3.0' for its calculations and inserts both the overall results for each impact category as well as the contribution trees.

How to use:

	1. Make sure you have an empty version (except the Parameters Check Tab) of the LCO2/Floating Model template excel file. Please note, that since we will be using the impact method 'EF 3.0' on the sheet 'LC stages EF' there should be 28 impact categories!
    2. Open the Database containg the LCO2/Floating Model.
    3. Run the script by clicking on the green button in the top left corner of your openLCA interface.
    4. A file selection window will open. Select the empty LCO2 model excel template file.
    5. Next a dialog box in which you can select from the set of product systems in your database. Select the product system 'LCO2 life cycle stage'.
    6. After this all the calculations will run. Please remain patient until the script has finished running! You will know when it has finished when a message in your console appears saying 'Finished! The file was saved under {FILE PATH}'.
    7. Now go to the directory under which you have saved the empty LCO2/Floating model excel template file. You should see a file with the dame file name only that the word '_FILLED' has been added at the end. This is the file containing the calculation results.

"""


from java.io import FileInputStream, FileOutputStream
from org.apache.poi.xssf.usermodel import XSSFWorkbook, XSSFFormulaEvaluator
from org.eclipse.jface.dialogs import MessageDialog
from org.openlca.app.components import FileChooser, ModelSelector
import string

def get_esg_product_system(product_system_uuid):
  """Get the ESG product system associated with the product system whose uuid was provided as an argument to this function."""
  
  selected_product_system = db.get(ProductSystem, product_system_uuid)
  split_name = selected_product_system.name.split()
  split_name.pop()
  split_name.append('ESG')
  esg_name = ' '.join(split_name)
  for product_system in db.getAll(ProductSystem):
    if product_system.name.startswith(esg_name):
      return product_system  

class Lco2Modeler():
  def __init__(self, excel_path, new_excel_path, product_system_uuid):
    self.excel_path = excel_path
    self.new_excel_path = new_excel_path
    self.product_system_uuid = product_system_uuid
    self.esg_product_system = get_esg_product_system(product_system_uuid)

  def update_parameter_set(self, warnings=False):
    """This function will update the parameters in the parameter set of the product system based on the the provided excel file and return that updated parameter set."""
    
    #get product system parameter set
    product_system = db.get(ProductSystem, self.product_system_uuid)
    parameter_set = product_system.parameterSets[0]
    
    #call workbook
    infile = FileInputStream(self.excel_path)
    workbook = XSSFWorkbook(infile)     
    
    #get 'Parameters check' sheet
    sheet = workbook.getSheet('Parameters check')
    
    #get column values for each column name in dictionary
    column_position_dict = {}
    header_row = sheet.getRow(0)
    for header_cell_num in range(header_row.lastCellNum + 1):
      try:
        header_cell = header_row.getCell(header_cell_num)
        column_position_dict[header_cell.getStringCellValue()] = header_cell_num
      except:
        if warnings:
          print('Cell does not exist!')
    
    #iterate through all parameters and update the parameter values
    for row_num in range(sheet.lastRowNum + 1):
      try:
        row = sheet.getRow(row_num)
        
        parameter_name_cell = row.getCell(column_position_dict['Parameter'])
        modified_value_cell = row.getCell(column_position_dict['Modified value'])
        if parameter_name_cell != None and modified_value_cell != None:
          parameter_name = parameter_name_cell.getStringCellValue()
          modified_value = modified_value_cell.getNumericCellValue()
          
          for parameter in parameter_set.parameters:
            if parameter.name == parameter_name:
              parameter.value = modified_value
              print("The parameter '%s' (uuid: %s) was updated to %f" % (parameter.name, parameter.refId, modified_value))
              counter += 1
      except:
        if warnings:
          print('Row does not exist!') 
    
    return parameter_set
    
  def run_impact_calculation(self, ef3_method=True, use_esg_product_system=False):
    """This function runs an impact calculation based on the product system and impact method whose uuid has been provided. If you set the parameter 'ef3_method' to True, the impact method 'EF 3.0 (adapted)' will be used to run the impact calculation. 
    If you set this parameter to False, the impact method 'Cumulative Energy Demand' will be used. If you want to use the 'normal' product system to make the calculations, leave the parameter 'use_esg_product_system' on False. If however, you wish to 
    make the calculation with the ESG equivalent, select 'use_esg_product_system' to True. A results object will be returned."""
	
    #get product system and impact method
    #get the ESG product system that is the equivalent of the product system indicated in the constructor
    if use_esg_product_system:
      product_system = self.esg_product_system
    else:
      product_system = db.get(ProductSystem, self.product_system_uuid)
    
    #get the impact method. Either 'EF 3.0 (adapted)' ('b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8') or 'Cumulative Energy Demand' ('be749018-2f47-3c25-819e-6e0c6fca1cb5') depending on the condition set when calling the function.
    if ef3_method:
      impact_method = db.get(ImpactMethod, 'b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8') 
    else:
      impact_method = db.get(ImpactMethod, 'be749018-2f47-3c25-819e-6e0c6fca1cb5')
    
    #create impact calculation setup
    calculator = SystemCalculator(db)
    setup = CalculationSetup.of(product_system)\
      .withImpactMethod(impact_method)\
      .withAllocation(AllocationMethod.NONE)\
      .withParameters(self.update_parameter_set().parameters)
      
    #run impact calculation
    return calculator.calculate(setup)
  
    
  def get_and_write_contribution_tree(self):
    """This function will return the contribution tree of the results object calculated using the impact method 'EF 3.0 Method (adapted)' with the function above. It will also write the contribution tree to a new tab in the excel file whose path was provided in the constructor."""
    
    #call workbook
    infile = FileInputStream(self.excel_path)
    workbook = XSSFWorkbook(infile)
    
    #run impact calculations
    results = self.run_impact_calculation()
    
    #create style
    style = workbook.createCellStyle()
    font = workbook.createFont()
    font.setBold(True)
    style.setFont(font)
    
    #delete existing 'Upstream tree' sheet if present
    sheet_names = [sheet.getSheetName() for sheet in workbook.sheetIterator()]
    for sheet_name in sheet_names:
      if 'Upstream tree' in sheet_name:
        sheet_id = sheet_names.index(sheet_name)
        workbook.removeSheetAt(sheet_id)
    
    #get a list of all the impact categories in the impact method 
    impact_categories = db.get(ImpactMethod, 'b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8').impactCategories
    
    #create an index sheet for the contribution tree sheets. The reason for this is that Excel has a character limit for the sheet names. Therefore, we cannot use the impact category names as the sheet names and instead we are using a numeric system for the sheet names that can be navigted using this index sheet.
    index_sheet = workbook.createSheet('Upstream tree sheet index')
    
    #create headers on index sheet
    index_header_row = index_sheet.createRow(0)
    sheet_name_header_cell = index_header_row.createCell(0)
    sheet_name_header_cell.setCellStyle(style)
    sheet_name_header_cell.setCellValue("Sheet Name")
    
    impact_category_header_cell = index_header_row.createCell(1)
    impact_category_header_cell.setCellStyle(style)
    impact_category_header_cell.setCellValue("Impact Cateory")
    
    #iterate and write index to index sheet
    for i, impact_category in enumerate(impact_categories, 1):
      index_row = index_sheet.createRow(i)
      
      index_sheet_name_cell = index_row.createCell(0)
      index_sheet_name_cell.setCellValue("Upstream tree %i" % i)
      
      index_impact_category_cell = index_row.createCell(1)
      index_impact_category_cell.setCellValue(impact_category.name)

    for index_column in range(2):
      index_sheet.autoSizeColumn(index_column)

    #iterate over impact categories in impact method
    for i, impact_category in enumerate(impact_categories, 1):
      
      #create sheet name
      sheet_name = "Upstream tree %i" % i
      
      #create sheet
      sheet = workbook.createSheet(sheet_name)
    
      #call reference process 
      reference_process = db.get(ProductSystem, self.product_system_uuid).referenceProcess
    
      #create and set cell A1
      row1 = sheet.createRow(0)
      cellA1 = row1.createCell(string.ascii_uppercase.index('A'))
      cellA1.setCellStyle(style)
      header_string = 'Upstream contributions to: %s' % impact_category.name
      cellA1.setCellValue(header_string)
      
      #create and set cell A2
      row2 = sheet.createRow(1)
      cellA2 = row2.createCell(string.ascii_uppercase.index('A'))
      cellA2.setCellStyle(style)
      cellA2.setCellValue('Processes')
      
      #create and set cell F2
      cellF2 = row2.createCell(string.ascii_uppercase.index('F'))
      results_header_string = 'Result [%s]' % impact_category.referenceUnit
      cellF2.setCellStyle(style)
      cellF2.setCellValue(results_header_string)
      
      #create and set cell G2
      cellG2 = row2.createCell(string.ascii_uppercase.index('G'))
      percentage_results_header_string = 'Percentage [%]'
      cellG2.setCellStyle(style)
      cellG2.setCellValue(percentage_results_header_string)      
      
      #build upstream contribution tree
      tech_flow = TechFlow.of(reference_process)
      tree = UpstreamTree.of(results.provider(), Descriptor.of(impact_category))
    
      #create upstream tree dictionary
      root_dict = {'name': tree.root.provider().provider().name, 'result': tree.root.result(), 'children': []}
      
      #get and write root level contribution
      print("%s: %s" % (tree.root.provider().provider().name, tree.root.result()))
      
      row3 = sheet.createRow(2)
      cellA3 = row3.createCell(string.ascii_uppercase.index('A'))
      cellA3.setCellValue(tree.root.provider().provider().name)
      
      cellF3 = row3.createCell(string.ascii_uppercase.index('F'))
      base_result = tree.root.result()
      cellF3.setCellValue(base_result)
      
      cellG3 = row3.createCell(string.ascii_uppercase.index('G'))
      root_percentage = tree.root.result() / base_result * 100
      cellG3.setCellValue(root_percentage)      
      
      counter = 3
      for child in tree.childs(tree.root):
        
        #get and write first level contribution
        child_dict = {'name': child.provider().provider().name, 'result': child.result(), 'children': []}
        print("\t%s: %s" % (child.provider().provider().name, child.result()))
        root_dict['children'].append(child_dict)
        
        rowx = sheet.createRow(counter)
        counter += 1
        cellBx = rowx.createCell(string.ascii_uppercase.index('B'))
        cellBx.setCellValue(child.provider().provider().name)
        cellFx = rowx.createCell(string.ascii_uppercase.index('F'))
        cellFx.setCellValue(child.result())
        cellGx = rowx.createCell(string.ascii_uppercase.index('G'))
        child_percentage = child.result() / base_result * 100
        cellGx.setCellValue(child_percentage)
        
        for grand_child in tree.childs(child):
          
          #get and write second level contribution
          grand_child_dict = {'name': grand_child.provider().provider().name, 'result': grand_child.result(), 'children': []}
          print("\t\t%s: %s" % (grand_child.provider().provider().name, grand_child.result()))
          child_dict['children'].append(grand_child_dict)
          
          rowx = sheet.createRow(counter)
          counter += 1
          cellCx = rowx.createCell(string.ascii_uppercase.index('C'))
          cellCx.setCellValue(grand_child.provider().provider().name)
          cellFx = rowx.createCell(string.ascii_uppercase.index('F'))
          cellFx.setCellValue(grand_child.result()) 
          cellGx = rowx.createCell(string.ascii_uppercase.index('G'))
          grand_child_percentage = grand_child.result() / base_result * 100
          cellGx.setCellValue(grand_child_percentage)
          
          for great_grand_child in tree.childs(grand_child):
            
            #get and write third level contribution
            great_grand_child_dict = {'name': great_grand_child.provider().provider().name, 'result': great_grand_child.result(), 'children': []}
            print("\t\t\t%s: %s" % (great_grand_child.provider().provider().name, great_grand_child.result()))
            grand_child_dict['children'].append(great_grand_child_dict)
  
            rowx = sheet.createRow(counter)
            counter += 1
            cellDx = rowx.createCell(string.ascii_uppercase.index('D'))
            cellDx.setCellValue(great_grand_child.provider().provider().name)
            cellFx = rowx.createCell(string.ascii_uppercase.index('F'))
            cellFx.setCellValue(great_grand_child.result())     
            cellGx = rowx.createCell(string.ascii_uppercase.index('G'))
            great_grand_child_percentage = great_grand_child.result() / base_result * 100
            cellGx.setCellValue(great_grand_child_percentage)
          
            for great_great_grand_child in tree.childs(great_grand_child):
              
              #get and write fourth level contribution
              great_great_grand_child_dict = {'name': great_great_grand_child.provider().provider().name, 'result': great_great_grand_child.result(), 'children': []}
              print("\t\t\t\t%s: %s" % (great_great_grand_child.provider().provider().name, great_great_grand_child.result()))
              great_grand_child_dict['children'].append(great_great_grand_child_dict)
      
              rowx = sheet.createRow(counter)
              counter += 1
              cellEx = rowx.createCell(string.ascii_uppercase.index('E'))
              cellEx.setCellValue(great_great_grand_child.provider().provider().name)
              cellFx = rowx.createCell(string.ascii_uppercase.index('F'))
              cellFx.setCellValue(great_great_grand_child.result()) 
              cellGx = rowx.createCell(string.ascii_uppercase.index('G'))
              great_great_grand_child_percentage = great_great_grand_child.result() / base_result * 100
              cellGx.setCellValue(great_great_grand_child_percentage)              
    
      #autoset the column size of all the columns on the created sheet
      for column in range(string.ascii_uppercase.index('G')+1):
        sheet.autoSizeColumn(column)
    
    #write all information to excel file
    outfile = FileOutputStream(self.new_excel_path)
    workbook.write(outfile)
    workbook.close()
    print('The upstream trees were written to excel file.')
    infile.close()
    outfile.close()    
    
    return root_dict
  
  def write_impact_calculation_results(self, warnings=False):
    """This function will write the calculation results using the impact method 'EF 3.0 Method (adapted)' to the excel file."""  
    
    #call workbook and sheet 
    infile = FileInputStream(self.new_excel_path)
    workbook = XSSFWorkbook(infile)
    sheet = workbook.getSheet('LC stages_EF')
        
    #get reference process to product system
    reference_process = db.get(ProductSystem, self.product_system_uuid).referenceProcess
    
    #get impact method
    impact_method = db.get(ImpactMethod, 'b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8')
    
    #run impact calculations
    results = self.run_impact_calculation()
    
    #iterate over all impact categories
    results_per_category_list = []
    for impact_category in impact_method.impactCategories:
      
      #build upstream contribution tree
      tech_flow = TechFlow.of(reference_process)
      tree = UpstreamTree.of(results.provider(), Descriptor.of(impact_category))
      for child in tree.childs(tree.root):
        
        #put all process names, results, impact categories and impact unit in a dictionary and save each of these dictionaries in a master list
        if child.provider().provider().name.endswith(' (float)'):
          child_dict = {'name': child.provider().provider().name.replace(' (float)', ''), 'result': child.result(), 'impact_category': impact_category.name, 'impact_unit': impact_category.referenceUnit}
        else:
          child_dict = {'name': child.provider().provider().name, 'result': child.result(), 'impact_category': impact_category.name, 'impact_unit': impact_category.referenceUnit}
        print "For the impact category '%s' in '%s' the impact is %f %s" % (child_dict['impact_category'], child_dict['name'], child_dict['result'], child_dict['impact_unit'])
        results_per_category_list.append(child_dict)
    
    #dictionary with the row location of each impact category on the sheet 'LC stages_EF'
    row_dict = {
      'Acidification': 1,
      'Climate change - Biogenic': 2,
      'Climate change - Fossil': 3,
      'Climate change - Land use and LU change': 4,
      'Climate change': 5,
      'Ecotoxicity, freshwater - inorganics': 6,
      'Ecotoxicity, freshwater - metals': 7,
      'Ecotoxicity, freshwater - organics': 8,
      'Ecotoxicity, freshwater': 9,
      'Eutrophication, freshwater': 10,
      'Eutrophication, marine': 11,
      'Eutrophication, terrestrial': 12,
      'Human toxicity, cancer - inorganics': 13,
      'Human toxicity, cancer - metals': 14,
      'Human toxicity, cancer - organics': 15,
      'Human toxicity, cancer': 16,
      'Human toxicity, non-cancer - inorganics': 17,
      'Human toxicity, non-cancer - metals': 18,
      'Human toxicity, non-cancer - organics': 19,
      'Human toxicity, non-cancer': 20,
      'Ionising radiation': 21,
      'Land use': 22,
      'Ozone depletion': 23,
      'Particulate matter': 24,
      'Photochemical ozone formation': 25,
      'Resource use, fossils': 26,
      'Resource use, minerals and metals': 27,
      'Water use': 28
    }
    
    #dictionary with the column location of each sub-process on the sheet 'LC stages_EF'
    column_dict = {
      'Material Extraction': string.ascii_uppercase.index('C'),
      'Manufacturing': string.ascii_uppercase.index('D'),
      'Operation': string.ascii_uppercase.index('G'),
      'Decommissioning': string.ascii_uppercase.index('H'),
      'Installation': string.ascii_uppercase.index('F'),
      'Transportation to Site': string.ascii_uppercase.index('E'), 
      'Site investigation': string.ascii_uppercase.index('B')
    }

    
    #iterate over results of impact categories and write them to the workbook
    for result_per_category in results_per_category_list:
      if result_per_category['impact_category'] in row_dict:
        row_num = row_dict[result_per_category['impact_category']]
        row = sheet.getRow(row_num)
        if result_per_category['name'] in column_dict: 
          cell_num = column_dict[result_per_category['name']]
          cell = row.getCell(cell_num)
          cell.setCellValue(result_per_category['result'])
    
    #evaluate all formulas in the workbook
    XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook)
    
    #write all information to excel file
    outfile = FileOutputStream(self.new_excel_path)
    workbook.write(outfile)
    print "Impact calculations have been written to the excel file."
    workbook.close()
    infile.close()
    outfile.close() 
    
  def write_main_components_results(self, warnings=False):
    """This function will write the contributions of the main components to the Excel file."""
    
    #call workbook and sheet 
    infile = FileInputStream(self.new_excel_path)
    workbook = XSSFWorkbook(infile)
    sheet = workbook.getSheet('CO2 eq. distribution')
        
    #get reference process to product system
    reference_process = db.get(ProductSystem, self.product_system_uuid).referenceProcess
    
    #get impact method
    impact_method = db.get(ImpactMethod, 'b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8')
    
    #run impact calculations
    results = self.run_impact_calculation()
    
    #iterate over all impact categories
    climate_change_results_dict = {}
    climate_change_impact_category = [impact_category for impact_category in impact_method.impactCategories if impact_category.name == 'Climate change'][0]
    
    #build upstream contribution tree
    tech_flow = TechFlow.of(reference_process)
    tree = UpstreamTree.of(results.provider(), Descriptor.of(climate_change_impact_category))
    for child in tree.childs(tree.root):
      for grand_child in tree.childs(child):
        if grand_child.provider().provider().name.endswith(' (float)') or grand_child.provider().provider().name.endswith(' (floating)'):
          climate_change_results_dict[grand_child.provider().provider().name.replace(' (float)', '').replace(' (floating)', '')] = grand_child.result()
        else:
          climate_change_results_dict[grand_child.provider().provider().name] = grand_child.result()
    print(climate_change_results_dict,'****')
    
    #iterate over rows in the main components table and insert the values
    for i in range(13, 25):
      row = sheet.getRow(i)
      name_cell = row.getCell(0)
      name = name_cell.getStringCellValue
      print(name,'++++name')
      value_cell = row.getCell(1)
      if name in climate_change_results_dict:
        value_cell.setCellValue(climate_change_results_dict[name])
        
    #evaluate all formulas in the workbook
    XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook)
    
    #write all information to excel file
    outfile = FileOutputStream(self.new_excel_path)
    print(outfile,'gsdghdm')
    workbook.write(outfile)
    print "Impact calculations for the main components have been written to the excel file."
    workbook.close()
    infile.close()
    outfile.close()     
          
  def write_esg_impact_calculation_results(self, warnings=False):
    """This function will write the ESG calculation results using the impact method 'EF 3.0 Method (adapted)' to the excel file."""  
    
    #call workbook and sheet 
    infile = FileInputStream(self.new_excel_path)
    workbook = XSSFWorkbook(infile)
    sheet = workbook.getSheet('LC stages_EF ESG')
    
    #get reference process to product system
    reference_process = self.esg_product_system.referenceProcess
    
    #get impact method
    impact_method = db.get(ImpactMethod, 'b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8')
    
    #run impact calculations
    results = self.run_impact_calculation(use_esg_product_system=True)
    
    #iterate over all impact categories
    results_per_category_list = []
    for impact_category in impact_method.impactCategories:
      
      #build upstream contribution tree
      tech_flow = TechFlow.of(reference_process)
      tree = UpstreamTree.of(results.provider(), Descriptor.of(impact_category))
      for child in tree.childs(tree.root):
        
        #put all process names, results, impact categories and impact unit in a dictionary and save each of these dictionaries in a master list
        if child.provider().provider().name.endswith(' (float)'):
          child_dict = {'name': child.provider().provider().name.replace(' (float)', ''), 'result': child.result(), 'impact_category': impact_category.name, 'impact_unit': impact_category.referenceUnit}
        else:
          child_dict = {'name': child.provider().provider().name, 'result': child.result(), 'impact_category': impact_category.name, 'impact_unit': impact_category.referenceUnit}
        print "For the impact category '%s' in '%s' the impact is %f %s" % (child_dict['impact_category'], child_dict['name'], child_dict['result'], child_dict['impact_unit'])
        results_per_category_list.append(child_dict)
    
    #dictionary with the row location of each impact category on the sheet 'LC stages_EF'
    row_dict = {
      'Acidification': 1,
      'Climate change - Biogenic': 2,
      'Climate change - Fossil': 3,
      'Climate change - Land use and LU change': 4,
      'Climate change': 5,
      'Ecotoxicity, freshwater - inorganics': 6,
      'Ecotoxicity, freshwater - metals': 7,
      'Ecotoxicity, freshwater - organics': 8,
      'Ecotoxicity, freshwater': 9,
      'Eutrophication, freshwater': 10,
      'Eutrophication, marine': 11,
      'Eutrophication, terrestrial': 12,
      'Human toxicity, cancer - inorganics': 13,
      'Human toxicity, cancer - metals': 14,
      'Human toxicity, cancer - organics': 15,
      'Human toxicity, cancer': 16,
      'Human toxicity, non-cancer - inorganics': 17,
      'Human toxicity, non-cancer - metals': 18,
      'Human toxicity, non-cancer - organics': 19,
      'Human toxicity, non-cancer': 20,
      'Ionising radiation': 21,
      'Land use': 22,
      'Ozone depletion': 23,
      'Particulate matter': 24,
      'Photochemical ozone formation': 25,
      'Resource use, fossils': 26,
      'Resource use, minerals and metals': 27,
      'Water use': 28
    }
    
    #dictionary with the column location of each sub-process on the sheet 'LC stages_EF'
    column_dict = {
      'Material Extraction': string.ascii_uppercase.index('C'),
      'Manufacturing': string.ascii_uppercase.index('D'),
      #'Operation': string.ascii_uppercase.index('G'),
      #'Decommissioning': string.ascii_uppercase.index('H'),
      'Installation': string.ascii_uppercase.index('F'),
      'Transportation to Site': string.ascii_uppercase.index('E'), 
      'Site investigation': string.ascii_uppercase.index('B')
    }
    
    #iterate over results of impact categories and write them to the workbook
    for result_per_category in results_per_category_list:
      if result_per_category['impact_category'] in row_dict:
        row_num = row_dict[result_per_category['impact_category']]
        row = sheet.getRow(row_num)
        if result_per_category['name'] in column_dict: 
          cell_num = column_dict[result_per_category['name']]
          cell = row.getCell(cell_num)
          cell.setCellValue(result_per_category['result'])
    
    #evaluate all formulas in the workbook
    XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook)
    
    #write all information to excel file
    outfile = FileOutputStream(self.new_excel_path)
    workbook.write(outfile)
    print "Impact calculations for the ESG have been written to the excel file."
    workbook.close()
    infile.close()
    outfile.close()     
    
  def write_cumulative_energy_demand_results(self):
    """This function will write the calculation results using the impact method 'Cumulative Energy Demand' to the excel file."""
    
    #call workbook and sheet 
    infile = FileInputStream(self.new_excel_path)
    workbook = XSSFWorkbook(infile)
    sheet = workbook.getSheet('Energy factors')    
      
    #run impact calculations for cumulative energy demand and save the results in a dictionary
    result = self.run_impact_calculation(ef3_method=False)   
    results_dict = {}
    for impact in result.getTotalImpacts():
      results_dict[impact.impact().name] = (impact.value(), impact.impact().referenceUnit)
      
    #iterate through impact result cells
    for i in range(2, 9):
      row = sheet.getRow(i)
      name_cell = row.getCell(0)
      value_cell = row.getCell(1)
      impact_name = name_cell.getStringCellValue()
      if impact_name in results_dict:
        value_cell.setCellValue(results_dict[impact_name][0])
        
    #evaluate all formulas in the workbook
    XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook)
    
    #write all information to excel file
    outfile = FileOutputStream(self.new_excel_path)
    workbook.write(outfile)
    print "The cummulative energy demand has been written to excel file."
    workbook.close()
    infile.close()
    outfile.close()       
    
def main():
  """This function will execute all the functions in the class 'Lco2Modeler' sequencially. Dialog windows will guide the user through the selection of the appropriate excel file, impact method and impact category."""
  
  #Select the Excel template file
  file = FileChooser.open('*.xlsx')
  if not file:
    MessageDialog.openError(None, "Error", "You must select an excel file")
    return
  
  excel_path = file.getAbsolutePath()
  new_excel_path = excel_path.replace('.xlsx', ' - FILLED.xlsx')
  
  #Select the product system
  product_system_descriptor = ModelSelector.select(ModelType.PRODUCT_SYSTEM)
  if not product_system_descriptor:
    MessageDialog.openError(None, "Error", "You must select a product system")
    return
  product_system_uuid = product_system_descriptor.refId
  
  l2m = Lco2Modeler(excel_path, new_excel_path, product_system_uuid)
  l2m.get_and_write_contribution_tree()
  l2m.write_impact_calculation_results()
  l2m.write_main_components_results()
  l2m.write_esg_impact_calculation_results()
  l2m.write_cumulative_energy_demand_results()
  
  print "Updated the entire excel file on path '%s'." % new_excel_path
          
if __name__ == '__main__':  
  
  App.runInUI("Writing Excel File...", main)


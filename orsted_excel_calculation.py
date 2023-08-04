"""
openLCA Version: 2.0.0

Description:

    This script will run the impact calculation for the 'LCO2 life cycle stages' model and insert the results into a
     LCO2 Model excel template file which you can select when starting the script. The script uses the impact method
     'EF 3.0' for its calculations and inserts both the overall results for each impact category and the contribution
     trees.

How to use:

    1. Make sure you have an empty version (except the Parameters Check Tab) of the LCO2/Floating Model template Excel
      file. Please note, that since we will be using the impact method 'EF 3.0' on the sheet 'LC stages EF' there should
      be 28 impact categories!
    2. Open the Database contain the LCO2/Floating Model.
    3. Run the script by clicking on the green button in the top left corner of your openLCA interface.
    4. A file selection window will open. Select the empty LCO2 model excel template file.
    5. Next a dialog box in which you can select from the set of product systems in your database. Select the product
     system 'LCO2 life cycle stage'.
    6. After this all the calculations will run. Please remain patient until the script has finished running! You will
      know when it has finished when a message in your console appears saying 'Finished! The file was saved under
      {FILE PATH}'.
    7. Now go to the directory under which you have saved the empty LCO2/Floating model excel template file. You should
      see a file with the dame file name only that the word '_FILLED' has been added at the end. This is the file
      containing the calculation results.

"""
from java.io import FileInputStream, FileOutputStream
from org.apache.poi.xssf.usermodel import XSSFWorkbook, XSSFFormulaEvaluator
from org.eclipse.jface.dialogs import MessageDialog
from org.openlca.app.components import FileChooser, ModelSelector
from org.openlca.core.results import UpstreamTree, UpstreamNode
from org.openlca.core.matrix.index import TechFlow
from org.openlca.core.model import ImpactCategory

# dictionary with the row location of each impact category on the sheet 'LC stages_EF'
STAGE_EF_ROWS = {
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
}  # type: dict[str, int]


def index_of(letter):  # type: (str) -> int
    return ord(letter) - 65


def letter_of(index):  # type: (int) -> str
    return chr(ord('@') + index + 1)


def context_of(param):  # type: ('ParameterRedef') -> str
    if param.contextId:
        process = db.get(Process, param.contextId)
        return process.name
    return "global"


class Path:

    def __init__(self, node, prefix=None):  # type: (UpstreamNode, Path) -> None
        self.prefix = prefix
        self.node = node
        self.length = 0 if prefix is None else 1 + prefix.length

    def append(self, node):  # type: (UpstreamNode) -> Path
        return Path(node, self)

    def count(self, tech_flow):  # type: (TechFlow) -> int
        c = 1 if tech_flow == self.node.provider() else 0
        return c + self.prefix.count(tech_flow) if self.prefix is not None else c


class UpstreamTreeSheet:
    MAX_DEPTH = 4
    MIN_CONTRIBUTION = 0.0

    def __init__(self, sheet, tree, impact_category):  # type: ('Sheet', UpstreamTree, ImpactCategory) -> None
        self.sheet = sheet
        self.tree = tree
        self.impact_category = impact_category

        self.total_result = None
        self.max_column = None
        self.row_index = None
        self.results = []

    def write_sheet(self):  # type: () -> None
        self.create_header()

        # write the tree
        self.row_index = 1
        self.max_column = 0
        self.total_result = self.tree.root.result()
        path = Path(self.tree.root)
        self.traverse(path)

        # write the results
        for i in range(len(self.results)):
            result = self.results[i]
            percentage = 100 * result / self.total_result if self.total_result != 0 else 0
            Excel.cell(self.sheet, i + 2, self.max_column + 1, result)
            Excel.cell(self.sheet, i + 2, self.max_column + 2, percentage)

    def traverse(self, path):  # type: (Path) -> None
        if self.row_index >= 1048574:  # is the maximum row number of an Excel sheet.
            return

        node = path.node
        result = path.node.result()
        if result == 0:
            return
        if self.total_result != 0:
            c = abs(result / self.total_result)
            if c < self.MIN_CONTRIBUTION:
                return

        self.write(path)

        if path.length > self.MAX_DEPTH - 1:
            return

        for child in self.tree.childs(node):
            self.traverse(path.append(child))

    def write(self, path):  # type: (Path) -> None
        self.row_index += 1
        self.results.append(path.node.result())
        col_index = path.length
        self.max_column = max(col_index, self.max_column)
        node = path.node
        if node.provider() is None or node.provider().provider() is None:
            return
        Excel.cell(self.sheet, self.row_index, col_index, node.provider().provider().name)

    def create_header(self):  # type: () -> None
        workbook = self.sheet.getWorkbook()
        Excel.cell(self.sheet, 0, index_of("A"), 'Upstream contributions to: %s' % self.impact_category.name)
        Excel.bold(workbook, self.sheet, 0, index_of("A"))

        Excel.cell(self.sheet, 1, index_of("A"), 'Processes')
        Excel.bold(workbook, self.sheet, 1, index_of("A"))

        Excel.cell(self.sheet, 1, index_of("F"), 'Result [%s]' % self.impact_category.referenceUnit)
        Excel.bold(workbook, self.sheet, 1, index_of("F"))

        Excel.cell(self.sheet, 1, index_of("G"), 'Percentage [%]')
        Excel.bold(workbook, self.sheet, 1, index_of("G"))


class Lco2Modeler:

    def __init__(self, source, target, system_id, warning=True):
        # type: (XSSFWorkbook, XSSFWorkbook, str, bool) -> None
        self.source = source
        self.target = target
        self.warning = warning

        self.system = db.get(ProductSystem, system_id)
        self.esg_system = self.get_esg_system(system_id)
        self.ef3_method = db.get(ImpactMethod, 'b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8')
        self.ced_method = db.get(ImpactMethod, 'be749018-2f47-3c25-819e-6e0c6fca1cb5')
        self.results = None

    @staticmethod
    def get_esg_system(system_id):  # type: (str) -> 'ProductSystem'
        """Get the ESG product system associated with the product system whose uuid was provided as an argument to this
        function."""
        selected_product_system = db.get(ProductSystem, system_id)
        split_name = selected_product_system.name.split()
        split_name.pop()
        split_name.append('ESG')
        esg_name = ' '.join(split_name)
        for product_system in db.getAll(ProductSystem):
            if product_system.name.startswith(esg_name):
                return product_system

    def set_systems_parameters(self):  # type: () -> None
        new_params = self.parse_workbook_parameters()
        for system in [self.system, self.esg_system]:
            self.set_system_parameters(system, new_params)

    @staticmethod
    def set_system_parameters(system, new_params):  # type: ('ProductSystem', dict[tuple[str, str], float]) -> None
        if not system.parameterSets:
            return
        print(new_params)
        for parameter_set in system.parameterSets:
            for param in parameter_set.parameters:
                modified_value = new_params.get((param.name, context_of(param)))
                print("updating " + param.name + " context: " + context_of(param) + " value: " + str(modified_value))
                if modified_value is not None:
                    param.value = modified_value
        db.update(system)

    def get_and_write_contribution_tree(self):  # type: () -> None
        """This function will return the contribution tree of the results object calculated using the impact method
        'EF 3.0 Method (adapted)' with the function above. It will also write the contribution tree to a new tab in
        the Excel file whose path was provided in the constructor."""

        results = self.get_result_no_arg()

        impact_categories = self.ef3_method.impactCategories

        self.delete_upstream_sheet()
        self.create_index_sheet(impact_categories)

        # iterate over impact categories in impact method
        for i, impact_category in enumerate(impact_categories, 1):
            self.create_upstream_sheet(i, impact_category, results)

        print('The upstream trees were written to excel file.')

    def write_impact_calculation_results(self):  # type: () -> None
        sheet = self.target.getSheet('LC stages_EF')
        result = self.get_result_no_arg()
        result_per_category_list = self.get_result_per_category(result)

        # dictionary with the column location of each sub-process on the sheet 'LC stages_EF'
        columns = {
            'Material Extraction': index_of('C'),
            'Manufacturing': index_of('D'),
            'Operation': index_of('G'),
            'Decommissioning': index_of('H'),
            'Installation': index_of('F'),
            'Transportation to Site': index_of('E'),
            'Site investigation': index_of('B')
        }

        self.write_results(sheet, result_per_category_list, STAGE_EF_ROWS, columns)
        print("Impact calculations have been written to the excel file.")

    def write_main_components_results(self):  # type: () -> None
        """This function will write the contributions of the main components to the Excel file."""
        sheet = self.target.getSheet('CO2 eq. distribution')

        results = self.get_result_no_arg()

        # iterate over all impact categories
        categories = self.ef3_method.impactCategories
        climate_change_category = [cat for cat in categories if cat.name == 'Climate change'][0]

        # build upstream contribution tree
        tree = UpstreamTree.of(results.provider(), Descriptor.of(climate_change_category))
        climate_change_results_dict = self.get_results_of(tree)

        # iterate over rows in the main components table and insert the values
        for i in range(13, 25):
            name_cell = Excel.cell(sheet, i, 0)
            if not name_cell.isPresent() or name_cell.get() is None:
                continue
            name = name_cell.get().getStringCellValue()
            value_cell = Excel.cell(sheet, i, 1)
            if not value_cell.isPresent() or value_cell.get() is None:
                continue
            value = climate_change_results_dict[name] if name in climate_change_results_dict else 0.0
            value_cell.get().setCellValue(value)

        # evaluate all formulas in the workbook
        XSSFFormulaEvaluator.evaluateAllFormulaCells(self.target)
        print("Impact calculations for the main components have been written to the excel file.")

    def write_esg_impact_calculation_results(self):  # type: () -> None
        """This function will write the ESG calculation results using the impact method 'EF 3.0 Method (adapted)' to
        the Excel file."""
        sheet = self.target.getSheet('LC stages_EF ESG')

        # run impact calculations
        results = self.run_impact_calculation(use_esg_system=True)
        results_per_category_list = self.get_result_per_category(results)

        # dictionary with the column location of each sub-process on the sheet 'LC stages_EF'
        columns = {
            'Material Extraction': index_of('C'),
            'Manufacturing': index_of('D'),
            # 'Operation': index_of('G'),
            # 'Decommissioning': index_of('H'),
            'Installation': index_of('F'),
            'Transportation to Site': index_of('E'),
            'Site investigation': index_of('B')
        }

        self.write_results(sheet, results_per_category_list, STAGE_EF_ROWS, columns)

        print("Impact calculations for the ESG have been written to the excel file.")

    def write_cumulative_energy_demand_results(self):  # type: () -> None
        """This function will write the calculation results using the impact method 'Cumulative Energy Demand' to the
        Excel file."""
        sheet = self.target.getSheet('Energy factors')

        # run impact calculations for cumulative energy demand and save the results in a dictionary
        result = self.run_impact_calculation(ef3_method=False)
        results_dict = {}
        for impact in result.getTotalImpacts():
            results_dict[impact.impact().name] = (impact.value(), impact.impact().referenceUnit)

        # iterate through impact result cells
        for i in range(2, 9):
            name_cell = Excel.cell(sheet, i, 0)
            if not name_cell.isPresent() or name_cell.get() is None:
                continue
            value_cell = Excel.cell(sheet, i, 1)
            if not value_cell.isPresent() or value_cell.get() is None:
                continue

            impact_name = name_cell.get().getStringCellValue()
            if impact_name in results_dict:
                value_cell.get().setCellValue(results_dict[impact_name][0])

        # evaluate all formulas in the workbook
        XSSFFormulaEvaluator.evaluateAllFormulaCells(self.target)
        print("The cumulative energy demand has been written to excel file.")

    def get_result_per_category(self, result):  # type: ('LcaResult') -> list[dict[str, 'Any']]
        results_per_category_list = []  # type: list[dict[str, 'Any']]
        for impact_category in self.ef3_method.impactCategories:
            # build upstream contribution tree
            tree = UpstreamTree.of(result.provider(), Descriptor.of(impact_category))
            for child in tree.childs(tree.root):
                results_per_category_list.append(self.get_info(child, impact_category))
        return results_per_category_list

    def write_results(self, sheet, results_per_category_list, rows, columns):
        # type: ('Sheet', list[dict[str, 'Any']], dict[str, int], dict[str, int]) -> None
        # iterate over results of impact categories and write them to the workbook
        for result_per_category in results_per_category_list:

            category = result_per_category['impact_category']  # type: str
            name = result_per_category['name']  # type: str
            result = result_per_category['result']  # type: float
            if category in rows and name in columns:
                row_index = rows[category]
                column = columns[name]
                Excel.cell(sheet, row_index, column, result)

        # evaluate all formulas in the workbook
        XSSFFormulaEvaluator.evaluateAllFormulaCells(self.target)

    @staticmethod
    def get_results_of(tree):  # type: (UpstreamTree) -> dict[str, float]
        results = {}
        for child in tree.childs(tree.root):
            for grand_child in tree.childs(child):
                _name = grand_child.provider().provider().name
                if _name.endswith(' (float)') or _name.endswith(' (floating)'):
                    name = _name.replace(' (float)', '').replace(' (floating)', '')
                    results[name] = grand_child.result()
                else:
                    results[_name] = grand_child.result()
        return results

    def get_result_no_arg(self):  # type: () -> 'LcaResult'
        if self.results is None:
            self.results = self.run_impact_calculation()
        return self.results

    @staticmethod
    def get_info(child, impact_category):  # type: (UpstreamNode, ImpactCategory) -> dict[str, 'Any']
        # put all process names, results, impact categories and impact unit in a dictionary
        _name = child.provider().provider().name
        name = _name if not _name.endswith(' (float)') else _name.replace(' (float)', '')

        child_dict = {
            'name': name,
            'result': child.result(),
            'impact_category': impact_category.name,
            'impact_unit': impact_category.referenceUnit
        }

        return child_dict

    def delete_upstream_sheet(self):  # type: () -> None
        # delete existing 'Upstream tree' sheet if present
        sheet_names = [sheet.getSheetName() for sheet in self.target.sheetIterator()]
        for sheet_name in sheet_names:
            if 'Upstream tree' in sheet_name:
                sheet_id = sheet_names.index(sheet_name)
                self.target.removeSheetAt(sheet_id)

    def create_index_sheet(self, impact_categories):  # type: (list[ImpactCategory]) -> None
        """
        create an index sheet for the contribution tree sheets. The reason for this is that Excel has a character limit
        for the sheet names. Therefore, we cannot use the impact category names as the sheet names, and instead we are
        using a numeric system for the sheet names that can be navigated using this index sheet.
        """
        sheet = self.target.createSheet('Upstream tree sheet index')

        Excel.cell(sheet, 0, 0, "Sheet Name")
        Excel.bold(self.target, sheet, 0, 0)

        Excel.cell(sheet, 0, 1, "Impact Category")
        Excel.bold(self.target, sheet, 0, 1)

        # iterate and write index to index sheet
        for i, impact_category in enumerate(impact_categories, 1):
            Excel.cell(sheet, i, 0, "Upstream tree %i" % i)
            Excel.cell(sheet, i, 1, impact_category.name)

        for index_column in range(2):
            sheet.autoSizeColumn(index_column)

    def run_impact_calculation(self, ef3_method=True, use_esg_system=False):  # type: (bool, bool) -> 'LcaResult'
        """
        This function runs an impact calculation based on the product system and impact method whose uuid has been
        provided. If you set the parameter 'ef3_method' to True, the impact method 'EF 3.0 (adapted)' will be used to
        run the impact calculation. If you set this parameter to False, the impact method 'Cumulative Energy Demand' 
        will be used. If you want to use the 'normal' product system to make the calculations, leave the parameter 
        'use_esg_product_system' on False. If however, you wish to make the calculation with the ESG equivalent, select
        'use_esg_product_system' to True. A results object will be returned.
        """

        # get product system and impact method
        # get the ESG product system that is the equivalent of the product system indicated in the constructor
        product_system = self.esg_system if use_esg_system else self.system

        # get the impact method. Either 'EF 3.0 (adapted)' ('b4571628-4b7b-3e4f-81b1-9a8cca6cb3f8') or 'Cumulative
        # Energy Demand' ('be749018-2f47-3c25-819e-6e0c6fca1cb5') depending on the condition set when calling the
        # function.
        impact_method = self.ef3_method if ef3_method else self.ced_method

        # create impact calculation setup
        calculator = SystemCalculator(db)
        setup = CalculationSetup.of(product_system) \
            .withImpactMethod(impact_method) \
            .withAllocation(AllocationMethod.NONE) \
            .withParameters(product_system.parameterSets[0].parameters)

        # run impact calculation
        return calculator.calculate(setup)

    def parse_workbook_parameters(self):  # type: () -> dict[tuple[str, str], float]
        """This function will update the parameters in the parameter set of the product system based on the provided
        Excel file and return that updated parameter set."""
        parameters = {}

        # get 'Parameters check' sheet
        sheet = self.source.getSheet('Parameters check')

        # get column values for each column name in dictionary
        columns = self.get_headers(sheet)

        # iterate through all the row and update the corresponding parameter.
        for row_num in range(1, sheet.lastRowNum + 1):
            parameters.update(self.parse_parameter(sheet, row_num, columns))

        return parameters

    def parse_parameter(self, sheet, row_index, columns):
        # type: ('Sheet', int, dict[str, int], 'ParameterRedefSet') -> dict[tuple[str, str], float]
        column_index = columns['Parameter']
        name_cell = Excel.cell(sheet, row_index, column_index)
        modified_value_cell = Excel.cell(sheet, row_index, columns['Modified value'])
        context_value_cell = Excel.cell(sheet, row_index, columns['Context'])

        if self.warning:
            if not (name_cell.isPresent() and modified_value_cell.isPresent() and context_value_cell.isPresent()):
                print("Fail to parse {column}{row} parameter.".format(column=letter_of(column_index), row=row_index))
                return {}
        if name_cell.get() is None or modified_value_cell.get() is None or context_value_cell.get() is None:
            print("A cell is empty on row {row}.".format(row=row_index))
            return {}

        try:
            name = name_cell.get().getStringCellValue()
            value = modified_value_cell.get().getNumericCellValue()
            context = context_value_cell.get().getStringCellValue()
            # avoid encoding issue by wrapping with str()
            return {(str(name), str(context)): value}
        except RuntimeException:
            print("Fail to parse {column}{row} parameter.".format(column=letter_of(column_index), row=row_index))

    def get_headers(self, sheet):  # type: ('Sheet') -> dict[str, int]
        column_position_dict = {}
        header_row = Excel.row(sheet, 0)
        for column_index in range(header_row.lastCellNum + 1):
            header_cell = Excel.cell(header_row, column_index)
            if header_cell.isPresent():
                column_position_dict[header_cell.get().getStringCellValue()] = column_index
            elif self.warning:
                column = letter_of(column_index)
                print('Cell {column}{row} does not exist!'.format(column=column, row=1))
        return column_position_dict

    def create_upstream_sheet(self, index, impact_category, results):
        # type: (int, ImpactCategory, 'LcaResult') -> None
        sheet_name = "Upstream tree %i" % index
        print('Creating {sheet_name}.'.format(sheet_name=sheet_name))
        sheet = self.target.createSheet(sheet_name)

        # build upstream contribution tree
        tree = UpstreamTree.of(results.provider(), Descriptor.of(impact_category))

        tree_sheet = UpstreamTreeSheet(sheet, tree, impact_category)
        tree_sheet.write_sheet()

        for column in range(index_of("G") + 1):
            sheet.autoSizeColumn(column)

    def write(self, path):  # type: (Path) -> None
        outfile = FileOutputStream(path)
        self.target.write(outfile)
        outfile.close()


def main():  # type: () -> None
    """This function will execute all the functions in the class 'Lco2Modeler' sequentially. Dialog windows will
    guide the user through the selection of the appropriate Excel file, impact method and impact category."""

    # Select the Excel template file
    file = FileChooser.open('*.xlsx')
    if not file:
        MessageDialog.openError(None, "Error", "You must select an excel file")
        return

    # Select the product system
    system_descriptor = ModelSelector.select(ModelType.PRODUCT_SYSTEM)
    if not system_descriptor:
        MessageDialog.openError(None, "Error", "You must select a product system")
        return
    system_id = system_descriptor.refId

    source_path = file.getAbsolutePath()
    source_fis = FileInputStream(source_path)
    source = XSSFWorkbook(source_fis)
    source_fis.close()

    target_fis = FileInputStream(source_path)
    target = XSSFWorkbook(target_fis)
    target_fis.close()

    l2m = Lco2Modeler(source, target, system_id)
    l2m.set_systems_parameters()
    l2m.get_and_write_contribution_tree()
    l2m.write_impact_calculation_results()
    l2m.write_main_components_results()
    l2m.write_esg_impact_calculation_results()
    l2m.write_cumulative_energy_demand_results()

    target_path = source_path.replace('.xlsx', ' - FILLED.xlsx')

    l2m.write(target_path)

    source.close()
    target.close()

    print("Updated the entire Excel file on path '%s'." % target_path)


if __name__ == '__main__':
    App.runInUI("Writing Excel File...", main)

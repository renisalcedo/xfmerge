import pandas as pd

class MergeExcel:
    def __init__(self, files, new_file='out'):
        self.files = files
        self.new_file = new_file

    def write_excel(self):
        """ Merges the data from the excel files and writes the new excel
        :type dframes: Pandas
        """
        dframes = self.get_excel_data()

        merged = pd.concat(dframes)
        merged.to_excel(f'{self.new_file}.xlsx')

    def get_excel_data(self):
        """ Grabs Excel data from files
        :rtype dframes: Pandas
        """
        excel_files = [pd.ExcelFile(name) for name in self.files]
        dframes = [d.parse(d.sheet_names[0], header=None, index_col=None) for d in excel_files]
        dframes[1:] = [df[1:] for df in dframes[1:]]

        return dframes

class ManageFiles:
    def __init__(self, filenames):
        self.filenames = filenames

    def get_file_names(self):
        """ Gets the file names and split by comma
        :rtype names: list[str]
        """
        names = self.filenames.split(',')
        names = [name.strip() for name in names]

        return names

    def get_ext(self):
        """ Returns the extension in the given files
        :rtype ext: str
        """
        names = self.get_file_names()
        return names[0].split('.')[1]

if __name__ == '__main__':
    """ INITIAL DATA INPUT """
    filenames = input('File names (comma separated): ')
    new_file = input('New file name (optional): ')

    """ PARSING THE INPUT DATA """
    manage_files = ManageFiles(filenames)
    filenames = manage_files.get_file_names()
    ext = manage_files.get_ext()

    """ INSTANTIATION AND LOGIC TO MERGE FILES """
    if ext == 'csv':
        print("CSV Not Implemented Yet")

    elif ext == 'xlsx':
        if new_file:
            merge_excel = MergeExcel(filenames, new_file)
        else:
            merge_excel = MergeExcel(filenames)

        merge_excel.write_excel()

    else:
        print('Unknown extension')
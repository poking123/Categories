import xlrd
import xlsxwriter
import os


#path = "C:/Users/brian.gao/Downloads/cFarm/categoryImport/All_Files"
path = "C:/Users/sale/Documents/ChemFarmImports/categoryImportExcel/categoryImportExcelInput/All_Files"
# goes through the levels
levelsLabel = []
levelsLabel.append("/Level_1")
levelsLabel.append("/Level_2")
levelsLabel.append("/Level_3")

# file names go here - level 1 in index 0, level 2 in index 1...
levels = []

# appends files (each level) to levels
for levelLabel in levelsLabel:
    files = os.listdir(path + levelLabel)
    level = []
    for file in files:
        level.append(file)
    levels.append(level)
    

# columns
category_id = []
c_ids = []
column = []
top = []
sort_order = []
multiparent = []
name = []
description = []
meta_title = []
meta_description = []
meta_keyword = []
SEO = []
children_library = []

# makes the allCategories excel file
allCategories = xlsxwriter.Workbook("../../categoryImportExcel/categoryImportExcelOutput/allCategoriesImport.xlsx")
allCatSheet = allCategories.add_worksheet("All Categories")

# puts all the headers in header
header = []
header.append("category_id")
header.append("column")
header.append("top")
header.append("sort_order")
header.append("multiparent")
header.append("name")
header.append("description")
header.append("meta_title")
header.append("meta_description")
header.append("meta_keyword")
header.append("SEO")
header.append("children_library")

# writes the header for the all categories excel files
for col in range(0,len(header)):
    allCatSheet.write(0,col,header[col])

# counts how many categories currently have been written
currentCount = 1;


# iterates through the levels and file_names to make an excel file for each
# category and adds to the big list
for i,level in enumerate(levels): 
    for file_name in level:
        # location of the excel file
        file_location = path + levelsLabel[i] + "/" + file_name
        # reads the current excel file
        workbook = xlrd.open_workbook(file_location)
        sheet = workbook.sheet_by_index(0)

        # gets all the columns for the current excel file
        index = 0
        temp_category_id = sheet.col_values(index,1)
        index = index + 1
        temp_column = sheet.col_values(index,1)
        index = index + 1
        temp_top = sheet.col_values(index,1)
        index = index + 1
        temp_sort_order = sheet.col_values(index,1)
        index = index + 1
        temp_multiparent = sheet.col_values(index,1)
        index = index + 1
        temp_name = sheet.col_values(index,1)
        index = index + 1
        temp_description = sheet.col_values(index,1)
        index = index + 1
        temp_meta_title = sheet.col_values(index,1)
        index = index + 1
        temp_meta_description = sheet.col_values(index,1)
        index = index + 1
        temp_meta_keyword = sheet.col_values(index,1)
        index = index + 1
        temp_SEO = sheet.col_values(index,1)
        index = index + 1
        temp_children_library = sheet.col_values(index,1)

        # adds on to the allCategories list
        category_id += temp_category_id
        column += temp_column
        top += temp_top
        sort_order += temp_sort_order
        multiparent += temp_multiparent
        name += temp_name
        description += temp_description
        meta_title += temp_meta_title
        meta_description += temp_meta_description
        meta_keyword += temp_meta_keyword
        SEO += temp_SEO
        children_library += temp_children_library


# write the combined data
for index in range(1,len(category_id)+1):
    i = index - 1
    allCatSheet.write(index,0,category_id[i])
    allCatSheet.write(index,1,column[i])
    allCatSheet.write(index,2,top[i])
    allCatSheet.write(index,3,sort_order[i])
    allCatSheet.write(index,4,multiparent[i])
    allCatSheet.write(index,5,name[i].strip())
    allCatSheet.write(index,6,description[i])
    allCatSheet.write(index,7,meta_title[i].strip())
    allCatSheet.write(index,8,meta_description[i].strip())
    allCatSheet.write(index,9,meta_keyword[i].strip())
    allCatSheet.write(index,10,SEO[i].strip())
    allCatSheet.write(index,11,children_library[i])

allCategories.close()


# CATEGORY ID LIBRARY
categoryIDLibrary = xlsxwriter.Workbook("../../categoryImportExcel/categoryImportExcelOutput/categoryIDLibrary.xlsx")
sheet1 = categoryIDLibrary.add_worksheet("Sheet 1")
# total number of files
total_num_files = len(name)

bgrd_color = categoryIDLibrary.add_format()
bgrd_color.set_bg_color('#bcf5bc')

# header
libraryHeader = []
libraryHeader.append("Category_Id")
libraryHeader.append("Category_Name")
libraryHeader.append("Parent(s)")
#libraryHeader.append("Children_Library")

# writes the header for the categoryIDLibrary excel files
for col in range(0,len(libraryHeader)):
    sheet1.write(0,col,libraryHeader[col],bgrd_color)

# column elements (matches header)
libraryElements = []
    
libraryElements.append(category_id)
libraryElements.append(name)
libraryElements.append(multiparent)
#libraryElements.append(children_library)

currentMultiparent = multiparent[0]
color = False

sheet1.write(1,0,category_id[0])
sheet1.write(1,1,name[0])
sheet1.write(1,2,multiparent[0])

for index in range(2,len(category_id)+1):
    i = index - 1
    
    if (currentMultiparent != multiparent[i]):
        currentMultiparent = multiparent[i]
        color = not color

    if (color):
        sheet1.write(index,0,category_id[i],bgrd_color)
        sheet1.write(index,1,name[i],bgrd_color)
        sheet1.write(index,2,multiparent[i],bgrd_color)
    else:
        sheet1.write(index,0,category_id[i])
        sheet1.write(index,1,name[i])
        sheet1.write(index,2,multiparent[i])


categoryIDLibrary.close()

print("Total Number Of Categories: " + str(total_num_files))




import xlrd
import xlsxwriter

levels = []
levels.append("CILevel1")
levels.append("CILevel2")
levels.append("CILevel3")

file_names = []
# ADD FILE NAMES HERE ACCORDING TO LEVEL
level1 = [] # LEVEL 1
level1.append("")

level2 = [] # LEVEL 2
level2.append("SignalingPathways")
level2.append("FunctionalPathways")
level2.append("BiosyntheticPathways")
level2.append("ResearchAreas")
level2.append("ProductTypes")

level3 = [] # LEVEL 3
level3.append("CytoskeletalRegulation")

# each level contains all the categories for that level
file_names.append(level1)
file_names.append(level2)
file_names.append(level3)

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
allCategories = xlsxwriter.Workbook("allCategories.xlsx")
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
for index,levelPrefix in enumerate(levels): #levelPrefix = "level_"
    for level in file_names[index]: #ex: level = "SignalingPathways"
        # location of the excel file
        file_name = levelPrefix + level
        #file_location = "C:/Users/sale/Documents/ChemFarmImports/categoryImportExcel/sheets/" + file_name + ".xlsx"

        file_location = "C:/Users/brian.gao/Downloads/cFarm/sheets/" + file_name + ".xlsx"

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

        #makes the excel file for the input excel file
        tempExcelFile = xlsxwriter.Workbook(levelPrefix+level+".xlsx")
        tempSheet = tempExcelFile.add_worksheet("Sheet 1")

        # writes the header for the all categories excel files
        for i in range(0,len(header)):
            tempSheet.write(0,i,header[i])

        #makes the array for counting the ids
        fakeIDs = []
        numCatInFile = len(temp_name)
        for i in range(currentCount,currentCount+numCatInFile):
            fakeIDs.append(i)
        currentCount += numCatInFile

        # write the combined data for the current excel file
        for tempIndex in range(1,numCatInFile+1):
            i = tempIndex - 1
            tempSheet.write(tempIndex,0,fakeIDs[i])
            tempSheet.write(tempIndex,1,temp_column[i])
            tempSheet.write(tempIndex,2,temp_top[i])
            tempSheet.write(tempIndex,3,temp_sort_order[i])
            tempSheet.write(tempIndex,4,temp_multiparent[i])
            tempSheet.write(tempIndex,5,temp_name[i].strip())
            tempSheet.write(tempIndex,6,temp_description[i].strip())
            tempSheet.write(tempIndex,7,temp_meta_title[i].strip())
            tempSheet.write(tempIndex,8,temp_meta_description[i].strip())
            tempSheet.write(tempIndex,9,temp_meta_keyword[i].strip())
            tempSheet.write(tempIndex,10,temp_SEO[i].strip())
            tempSheet.write(tempIndex,11,temp_children_library[i])

        tempExcelFile.close()
        print(levelPrefix+level+" Saved.")

        # adds on to the big list of all of the columns
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


#makes category_id list
fakeCatID = []
numTotalCategories = len(name)
for i in range(1,numTotalCategories+1):
    fakeCatID.append(i)

# write the combined data
for index in range(1,numTotalCategories+1):
    i = index - 1
    allCatSheet.write(index,0,fakeCatID[i])
    allCatSheet.write(index,1,column[i])
    allCatSheet.write(index,2,top[i])
    allCatSheet.write(index,3,sort_order[i])
    allCatSheet.write(index,4,multiparent[i])
    allCatSheet.write(index,5,name[i].strip())
    allCatSheet.write(index,6,description[i].strip())
    allCatSheet.write(index,7,meta_title[i].strip())
    allCatSheet.write(index,8,meta_description[i].strip())
    allCatSheet.write(index,9,meta_keyword[i].strip())
    allCatSheet.write(index,10,SEO[i].strip())
    allCatSheet.write(index,11,children_library[i])

allCategories.close()


# CATEGORY ID LIBRARY
categoryIDLibrary = xlsxwriter.Workbook("categoryIDLibrary.xlsx")
sheet1 = categoryIDLibrary.add_worksheet("Sheet 1")
# total number of files
total_num_files = len(name)

print("Total Number Of Categories: " + str(total_num_files))
columns = total_num_files // 100
remainder = total_num_files % 100

# header
libraryHeader = []
libraryHeader.append("Category_Id")
libraryHeader.append("Category_Name")
libraryHeader.append("Parent(s)")
libraryHeader.append("Children_Library")

# column elements (matches header)
libraryElements = []
for i in range(1,len(category_id)+1): # list from 1 to total number of categories
    c_ids.append(i)
    
libraryElements.append(c_ids)
libraryElements.append(name)
libraryElements.append(multiparent)
libraryElements.append(children_library)

# if remainder, make column columns and then remainder column

# makes column columns
for i in range(0,columns): 
    for j in range(0,len(libraryHeader)): # 4 headers
        # writes header
        sheet1.write(0,5*i+j,libraryHeader[j])
        # writes data in that column
        for k in range(0,100):
            sheet1.write(k+1,5*i+j,libraryElements[j][100*i+k])

# makes remainder column
for j in range(0,len(libraryHeader)): # 4 headers
        # writes header
        sheet1.write(0,5*columns+j,libraryHeader[j])
        # writes data in that column
        for k in range(0,remainder):
            sheet1.write(k+1,5*columns+j,libraryElements[j][100*columns+k])

categoryIDLibrary.close()
    




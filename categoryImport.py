import xlrd
import codecs

file_name = "allCategories"
# make sure the path to the excel file you are reading from is correct
#file_location = "C:/Users/sale/Documents/ChemFarmImports/categoryImportExcel/" + file_name + ".xlsx"
file_location ="C:/Users/sale/Documents/ChemFarmImports/categoryImportSQL/Python/categoryImportExcelOutput/" + file_name + ".xlsx"

workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_index(0)

index = 0
category_id = sheet.col_values(index,1)
index = index + 1
column = sheet.col_values(index,1)
index = index + 1
top = sheet.col_values(index,1)
index = index + 1
sort_order = sheet.col_values(index,1)
index = index + 1
multiparent = sheet.col_values(index,1)
index = index + 1
name = sheet.col_values(index,1)
index = index + 1
description = sheet.col_values(index,1)
index = index + 1
meta_title = sheet.col_values(index,1)
index = index + 1
meta_description = sheet.col_values(index,1)
index = index + 1
meta_keyword = sheet.col_values(index,1)
index = index + 1
SEO = sheet.col_values(index,1)
index = index + 1
children_library = sheet.col_values(index,1)




# delete
# for some reason, importing into phpmyadmin cuts off the first 2 letters
# so I added in an extra "DE" below so it is correct
delete_oc_category = "DEDELETE FROM `oc_category` WHERE "
delete_oc_category_multiparent = "DELETE FROM `oc_category_multiparent` WHERE "
delete_oc_category_description = "DELETE FROM `oc_category_description` WHERE "
delete_oc_category_path = "DELETE FROM `oc_category_path` WHERE "
delete_oc_category_to_store = "DELETE FROM `oc_category_to_store` WHERE "
delete_oc_category_to_layout = "DELETE FROM `oc_category_to_layout` WHERE "
delete_oc_url_alias = "DELETE FROM `oc_url_alias` WHERE "

# insert
oc_category = "INSERT INTO `oc_category` (`category_id`,`column`,`top`,`sort_order`,`status`,`children_library`,`date_added`, `date_modified`) VALUES\n"
oc_category_multiparent = "INSERT INTO `oc_category_multiparent` (`category_id`,`parent_id`) VALUES\n"
oc_category_description = "INSERT INTO `oc_category_description` (`category_id`,`language_id`,`name`,`description`,`meta_title`,`meta_description`,`meta_keyword`) VALUES\n"
oc_category_path = "INSERT INTO `oc_category_path` (`category_id`,`path_id`,`level`) VALUES\n"
oc_category_to_store = "INSERT INTO `oc_category_to_store` (`category_id`,`store_id`) VALUES\n"
oc_category_to_layout = "INSERT INTO `oc_category_to_layout` (`category_id`,`store_id`,`layout_id`) VALUES\n"
oc_url_alias = "INSERT INTO `oc_url_alias` (`query`,`keyword`) VALUES\n"

# ITERATIONS
for i in range(len(category_id)-1):
    # delete
    delete_oc_category += "category_id='" + str(category_id[i]) + "' OR "
    delete_oc_category_multiparent += "category_id='" + str(category_id[i]) + "' OR "
    delete_oc_category_description += "category_id='" + str(category_id[i]) + "' OR "
    delete_oc_category_path += "category_id='" + str(category_id[i]) + "' OR "
    delete_oc_category_to_store += "category_id='" + str(category_id[i]) + "' OR "
    delete_oc_category_to_layout += "category_id='" + str(category_id[i]) + "' OR "
    delete_oc_url_alias += "query='category_id=" + str(category_id[i]) + "' OR "

    # oc_category
    oc_category += "('" + str(category_id[i]) + "','" + str(column[i]) + "','" + str(top[i]) + "','" + str(sort_order[i]) + "',1,'" + str(children_library[i]) + "',NOW(),NOW()),\n"

    # oc_category_multiparent
    if (multiparent[i]):
        if (not isinstance(multiparent[i],float)):
            parent_categories = multiparent[i].split(",")
            for j in range(len(parent_categories)):
                oc_category_multiparent += "('" + str(category_id[i]) + "','" + str(parent_categories[j]) + "'),\n"
        else:
            oc_category_multiparent += "('" + str(category_id[i]) + "','" + str(multiparent[i]) + "'),\n"
    

    # oc_category_description
    oc_category_description += "('" + str(category_id[i]) + "',1,'" + str(name[i]) + "','" + str(description[i]) + "','" + str(meta_title[i]) + "','" + str(meta_description[i]) + "','" + str(meta_keyword[i]) + "'),\n"
                                         
    # oc_category_path
    oc_category_path += "('" + str(category_id[i]) + "','" + str(category_id[i]) + "',0),\n"

    # oc_category_to_store
    oc_category_to_store += "('" + str(category_id[i]) + "',0),\n"                                    

    # oc_category_to_layout
    oc_category_to_layout += "('" + str(category_id[i]) + "',0,0),\n"

    # oc_url_alias
    oc_url_alias += "('category_id=" + str(category_id[i]) + "','" + str(SEO[i]) + "'),\n"                              
    
# after for loop (LAST ITERATION)
i = i + 1

# delete
delete_oc_category += "category_id='" + str(category_id[i]) + "';\n"
delete_oc_category_multiparent += "category_id='" + str(category_id[i]) + "';\n"
delete_oc_category_description += "category_id='" + str(category_id[i]) + "';\n"
delete_oc_category_path += "category_id='" + str(category_id[i]) + "';\n"
delete_oc_category_to_store += "category_id='" + str(category_id[i]) + "';\n"
delete_oc_category_to_layout += "category_id='" + str(category_id[i]) + "';\n"
delete_oc_url_alias += "query='category_id=" + str(category_id[i]) + "';\n"    


# oc_category
oc_category += "('" + str(category_id[i]) + "','" + str(column[i]) + "','" + str(top[i]) + "','" + str(sort_order[i]) + "',1,'" + str(children_library[i]) + "',NOW(),NOW());\n\n"

# oc_category_multiparent
if (multiparent[i]):
    if (not isinstance(multiparent[i],float)):
        parent_categories = multiparent[i].split(",")
        for j in range(len(parent_categories)-1):
                oc_category_multiparent += "('" + str(category_id[i]) + "','" + parent_categories[j] + "'),\n"
        j = j + 1
        oc_category_multiparent += "('" + str(category_id[i]) + "','" + str(parent_categories[j]) + "');\n\n"
    else:
        oc_category_multiparent += "('" + str(category_id[i]) + "','" + str(multiparent[i]) + "');\n\n"
    
# oc_category_description
oc_category_description += "('" + str(category_id[i]) + "',1,'" + str(name[i]) + "','" + str(description[i]) + "','" + str(meta_title[i]) + "','" + str(meta_description[i]) + "','" + str(meta_keyword[i]) + "');\n\n"                                         

# oc_category_path
oc_category_path += "('" + str(category_id[i]) + "','" + str(category_id[i]) + "',0);\n\n"

# oc_category_to_store
oc_category_to_store += "('" + str(category_id[i]) + "',0);\n\n"

# oc_category_to_layout
oc_category_to_layout += "('" + str(category_id[i]) + "',0,0);\n\n"                                     

# oc_url_alias
oc_url_alias += "('category_id=" + str(category_id[i]) + "','" + str(SEO[i]) + "');\n\n" 

# writing
with codecs.open("../categoryImportSQLOutput/"+file_name + ".sql","w","utf-8-sig") as temp:
    
    # deleting
    temp.write(delete_oc_category)
    temp.write(delete_oc_category_multiparent)
    temp.write(delete_oc_category_description)
    temp.write(delete_oc_category_path)
    temp.write(delete_oc_category_to_store)
    temp.write(delete_oc_category_to_layout)
    temp.write(delete_oc_url_alias)

    temp.write("\n")
    # adding
    temp.write(oc_category)
    if (oc_category_multiparent != "INSERT INTO `oc_category_multiparent` (`category_id`,`parent_id`) VALUES\n"):
        temp.write(oc_category_multiparent)
    temp.write(oc_category_description)
    temp.write(oc_category_path)
    temp.write(oc_category_to_store)
    temp.write(oc_category_to_layout)
    temp.write(oc_url_alias)


    temp.close()

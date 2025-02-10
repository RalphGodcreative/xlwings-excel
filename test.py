import xlwings as xw

book = xw.Book(r'C:/Users/user/Documents/insurance excel/example.xls')

# for finding the index of the print_sheet

# Loop through the sheets and find the index
def get_sheet_index(book, sheet_name):

    for index, sheet in enumerate(book.sheets):
        if sheet.name == sheet_name:
            return index
    return -1
    
    
input_sheet = book.sheets[get_sheet_index(book, "輸入頁")]
print_sheet = book.sheets[get_sheet_index(book, "列印頁")]
year_type = input_sheet.range('S5')
born_year = input_sheet.range('F7')
age = input_sheet.range('F10')
price = input_sheet.range('S8')
gender = input_sheet.range('F6')
year_type.value ="二十年期"
price.value = 100
all_gender = ("男","女")

all = list()

# output_columns = list(range(13,112))
output_columns = list(range(12,38)) + list(range(44,79)) + list(range(85, 120)) + list(range(126, 141))
output_list = ('AD' , 'AA')

for g in all_gender:
    gender.value = g    
    for i in range(54,115):
        
        # born_year.value = f"{i}0101"
        born_year.value=i
        print(f'{born_year.value} 生')
        print(age.value)
        # every line by year
        per_year= []
        for li in output_list:
            per_year.append(year_type.value)
            per_year.append(gender.value)
            per_year.append(age.value)
            for j in output_columns:
                per_year.append(print_sheet.range(li+str(j)).value)
            per_year.append('next-line')    
        per_year.pop()    
        all.append(per_year)
        print(per_year)       
    
book.close()

def next_column(alpha):
    # Convert the column string to a "number" (e.g., 'A' -> 1, 'B' -> 2, ..., 'Z' -> 26)
    result = []
    carry = 1  # Start with a carry to increment the column

    for char in reversed(alpha):
        new_char = ord(char) + carry
        if new_char > ord('Z'):  # If it goes beyond 'Z', wrap around to 'A' and keep the carry
            new_char = ord('A')
            carry = 1
        else:
            carry = 0  # No carry needed once we handle it
        result.append(chr(new_char))

    # If there's still a carry after processing all characters, add a new 'A' at the beginning
    if carry:
        result.append('A')

    # Reverse the result to get the final column name in the correct order
    return ''.join(reversed(result))

#where to start the ouput
column = 'D'

out_book = xw.Book(r'C:/Users/user/Documents/insurance excel/output/output-example.xlsx')
out_sheet = out_book.sheets[1]

for i in  range(len(all)):
    k=1
    for j in range(len(all[i])):
        if all[i][j]=='next-line':
            k = 1
            column = next_column(column)
        else:    
            out_sheet.range(column+str(k)).value = all[i][j]
            k=k+1
    column = next_column(column)     
            
# out_book.save()    
# out_book.close()



        
        
    
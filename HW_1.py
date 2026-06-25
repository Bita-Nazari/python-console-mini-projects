import random
import os
from colorama import Fore, Style,Back ,init
init()

import pygame
soundplay =pygame.mixer.init()

import win32com.client
speaker = win32com.client.Dispatch("SAPI.SpVoice")
speaker.Rate = 1

book_dicts = {} 
product_dicts = {}

#region setting

#region SoundSetting
def play_succes_sound():
    succes_sound =pygame.mixer.Sound('SuccesSound.mp3')
    succes_sound.play()
    
def play_changeSuccess_sound():
    succes_sound =pygame.mixer.Sound('SuccesAddItem.mp3')
    succes_sound.play()
    
def play_warning_sound():
    succes_sound =pygame.mixer.Sound('WarningSound.mp3')
    succes_sound.play()
    

def play_Fail_sound():
    succes_sound =pygame.mixer.Sound('FailSound.mp3')
    succes_sound.play()
#endregion            
    

#region PrintSetting
def convert_text_to_audio(message):
    

    speaker.Speak(message)


def print_warning_messages(message):

    
    print(Back.YELLOW + Fore.WHITE + message +Style.RESET_ALL)
        
def print_success_messages(message):

    print(Back.GREEN + Fore.WHITE + message +Style.RESET_ALL)  
    
def print_error_messages(message):

    print(Back.RED + Fore.RED + message +Style.RESET_ALL) 

def print_normal_messages(message):

    print(Back.MAGENTA + Fore.WHITE + message +Style.RESET_ALL) 

def is_input_in_range(range_number, user_input):
    range_of_numbers = range(1, range_number+1)
    if user_input not in range_of_numbers:
        play_warning_sound()
        print_warning_messages('The Input Number Is Not Defined In Program Please Enter The Correct Number')
        return False
    return True
#endregion

#endregion

# region random_number_game


def random_number_game():
    '''Its The Function For Playing The RandomNumber Game'''
    random_number = random.randint(0, 100)
    user_chance = 0
    while user_chance < 5:
        remain_chances = 5 - user_chance
        is_in_format = False
        while not is_in_format:
            try:
                user_guess = int(
                input(f'Enter Your Guess From 0 To 100 (Remain Chances : {remain_chances}) : '))
                is_in_format = True
            except:
                play_warning_sound()
                
                print_error_messages('Please Enter Digit Number')
            
            


        
        if user_guess < 0 or user_guess > 100:
            print_error_messages('Please Enter a Number Between 0 To 100')
        elif (random_number == user_guess):
            msg =f'Correct!The Number is {random_number} , You Did it in {user_chance + 1}th Guess'
            
            play_succes_sound()
            print_success_messages(msg)
            
            
                
            break
        else:
            user_chance = user_chance + 1
            if random_number > user_guess:
                if user_chance < 5:
                    msg =f'Wrong!Try Again (Guess a Bigger Number)'
                    
                    print_warning_messages(msg)
                    convert_text_to_audio(msg)

            else:
                if user_chance < 5:
                    msg =f'Wrong!Try Again (Guess a Smaller Number)'
                    
                    print_warning_messages(msg)
                    convert_text_to_audio(msg)
                    
    if user_chance == 5:
        msg =f'GameOver!!!The Number Was {random_number}'
        
        play_Fail_sound()
        print_error_messages(msg)
        


# endregion

# region Library Functions


def add_book(book_dicts):
    book_name = input('Enter Name Of The Book : ')
    author_name = input('Enter Name Of The Author : ')

    if book_name in book_dicts:
        msg =f'There is Already a Book in Library Named : {book_name}'
        
        print_warning_messages(msg)
        convert_text_to_audio(msg)
    else:
        book_dicts[book_name] = author_name
        play_changeSuccess_sound()
        print_success_messages('Your Book Added Succesfully :))')


def search_book(book_dicts):
    book_name = input('Enter Name Of The Book That You Want To Find : ')
    author_name = ''

    if book_name not in book_dicts:
        msg =(f'There Is No Book Such as {book_name} in The Library!')
        
        print_warning_messages(msg)
        convert_text_to_audio(msg)
        

    else:
        author_name = book_dicts[book_name]
        msg =(f'The Author Of The {book_name} is {author_name}')
        
        print_normal_messages(msg)
        convert_text_to_audio(msg)
       
def show_book(book_dicts):
    books= ''
    i=0
    for item in book_dicts:
        Author = book_dicts[item]
        books = books + '\n' + f'{i + 1} - Book`s Name : {item} - Author`s Name : {Author}'
        i +=1

    print(books)




def library_managment(book_dicts):
    msg = """
 << Library Menu >>
 1-Add
 2-Search
 3-Show
 4-Back To The Menu : """
    

    while True:
        
        is_check = False
        while is_check == False:
            
            is_in_format = False
            while not is_in_format:
                try:
                    user_input = int(
                    input(f'{msg} : '))
                    is_in_format = True
                except:
                    play_warning_sound()
                    print_error_messages('Please Enter Digit Number')
                is_check = is_input_in_range(4, user_input)

        if user_input == 4:
            break

        elif user_input == 1:
            add_book(book_dicts)

        elif user_input == 2:
            search_book(book_dicts)
        else:
            show_book(book_dicts)

# endregion

# region Inventory Functions

def read_product_files():
    with open('Products.txt' , 'r') as file:
        file_content = file.read()
        if len(file_content) > 0:
            lines = file_content.split('\n')
            return lines
        
    

def get_ready_product_dicts(product_dicts):
    file_content = read_product_files()
    if(len(product_dicts)  == 0 ):
        if len(file_content) > 0:
            for line in file_content:
                items = line.strip().split('-')
                key_item = items[0].strip()
                value_item = items[1].strip()
                product_dicts[key_item] = value_item
                
    return product_dicts
        

            
    


def is_product_by_dict(product_name ):
    is_product = True
    if not product_name in product_dicts:
        play_warning_sound()
        msg =(f'There is No {product_name} in Products')
        
        print_warning_messages(msg)
        convert_text_to_audio(msg)
        is_product = False
    return is_product

def is_product_by_file(product_name ):
    is_product = True
    file_line_content = read_product_files()
    file_product_names =[]
    for line in file_line_content:
        items = line.strip().split('-')
        key_item = items[0].strip()
        file_product_names.append(key_item)
        
    if not product_name in file_product_names:
        play_warning_sound()
        msg =print(f'There is No {product_name} in Products')
        
        print_warning_messages(msg)
        convert_text_to_audio(msg)
        is_product = False
        
    return is_product

def add_product(product_dicts):
    product_dicts = get_ready_product_dicts(product_dicts)
    product_name = input('Enter Name Of The Product : ')

    try:
        product_stock = int(input('Enter Count Of The Stock : '))
    except:
        play_warning_sound()
        print_error_messages('Please Enter The Input in Digit Format')
 
    if product_name in product_dicts:
        old_stock = product_dicts[product_name]
        product_dicts[product_name] = int(old_stock )+ product_stock

    else:
        product_dicts[product_name] = product_stock
        play_changeSuccess_sound()
        print_success_messages('Your Product Added Succesfully(For Save New Changes in File Do Action Save)')

def sell_product(product_dicts):
    product_dicts = get_ready_product_dicts(product_dicts)
    product_name = input('Enter Name Of The Product : ')
    
    is_in_format = False
    while not is_in_format:
        try:
            sell_count = int(input('Enter Count Of The Sell : '))
            is_in_format = True
        except:
            play_warning_sound()
            is_in_format = False
            print_error_messages('Please Enter The Input in Digit Format')
        
        
    is_product = is_product_by_dict(product_name)
    if is_product == False:
        return

    stock = product_dicts[product_name]
    if sell_count > stock:
        msg = (f'There is Not Enough Stock of {product_name} in Inventory (There is Only {stock})')
        print_error_messages(msg)
        convert_text_to_audio(msg)
        return
    product_dicts[product_name] = stock - sell_count
    play_changeSuccess_sound()
    print_success_messages(f'Your Product Selled Succesfully!!')

        

def search_product():
    file_line_content  =read_product_files()
    product_name = input('Enter Name Of The Product You Want To Find: ')
    is_product = is_product_by_file(product_name)
    if is_product == False:
        return
    
    for line in file_line_content:
        items = line.strip().split('-')
        if(items[0].strip() == product_name):
            msg =(f'The Stock Of {product_name} is {items[1]}')
            
            print_normal_messages(msg)
            convert_text_to_audio(msg)
            
            break
    



def show_products():
    file_line_content = read_product_files()
    products = ''
    i=0
    for line in file_line_content:
        items = line.strip().split('-')
        product_name = items[0]
        product_stock = items[1]
        products = products + '\n' + f'{i + 1} - Products`s Name : {product_name} - Product`s St : {product_stock}'
        i +=1

    print(products)

def save_product_in_file(product_dicts):
    product_dicts = get_ready_product_dicts(product_dicts)
    file_name = 'Products.txt'
    

    if not os.path.exists(f'.\{file_name}'):
            with open(file_name,'a') as f:
                f.close
                

        
    with open(file_name , 'w') as file:
        lines = []
        for item in product_dicts:
            line = f'{item} - {product_dicts[item]}'
            lines.append(line)
            
        file.write('\n'.join(lines))
        play_changeSuccess_sound()
        msg =('Your Changes Saved SuccesFully')
        
        print_success_messages(msg)
        convert_text_to_audio(msg)


def get_inventory_report():
    
    file_line_content =  read_product_files()
    i = 0
    sum_of_product_stocks =0
    products_have_stock =0


    for line in file_line_content:
        
        items = line.strip().split('-')
        if i == 0 :
            min_products_stocks = items[0]
            min_products_stocks_count=int(items[1]) 
            
            max_of_products_stocks = items[0]
            max_of_products_stocks_count =int(items[1]) 


        

        if int(items[1])  > 0:
            products_have_stock += 1
        sum_of_product_stocks +=  int(items[1])
        if int(items[1]) > max_of_products_stocks_count:
            max_of_products_stocks = items[0]
            max_of_products_stocks_count = int(items[1])
        if int(items[1]) < min_products_stocks_count:
            min_products_stocks = items[0]
            min_products_stocks_count = int(items[1])
        i +=1
    
    print(f'Products Have Stock : {products_have_stock}\nSum Of Stocks : {sum_of_product_stocks}\nProduct With Max Stock : {max_of_products_stocks}\nProducts With Min Stock : {min_products_stocks}')






def inventory_managment(product_dicts):
    msg = """
 << Inventory Menu >>
 1-Add
 2-Sell
 3-Search
 4-Show
 5-Save
 6-Report
 7-Back To The Menu : """
 
    warning_msg ="""as long as you don't save your changes to the file, you won't see them in the file."""
    while True:
            is_check = False
            while is_check == False:
                user_input = int(input(msg))
                is_check = is_input_in_range(7, user_input)

            if user_input == 7:
                break

            elif user_input == 1:
                add_product(product_dicts)

            elif user_input == 2:
                sell_product(product_dicts)
                
            
            elif user_input ==3 :
                print_warning_messages(warning_msg)
                search_product()
                
            
            elif user_input == 4:
                print_warning_messages(warning_msg)
                show_products()
                
            
            elif user_input == 5:
                save_product_in_file(product_dicts)
                
            elif user_input == 6:
                print_warning_messages(warning_msg)
                get_inventory_report()
                



# endregion
msg = """
Welcome
1-RandomNumber Game
2-Manage Library
3-Manage Markets's Inventory
(press 4 to exit)
Enter Number Of The Question : """
is_running = True

while is_running:
    
    is_check = False
    while is_check == False:
        q_choice = int(input(msg))
        is_check = is_input_in_range(4, q_choice)
    if q_choice == 1:
        random_number_game()

    elif q_choice == 2:
        library_managment(book_dicts)
    elif q_choice == 3:
        inventory_managment(product_dicts)
    elif q_choice == 4:
        break


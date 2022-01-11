'''write a programe which will keep adding a strea of numbers inputted by the user . The adding stops as soon as user passes q on the keyborad'''
sum = 0
while(True):
    userInput = input("Enter the item price or press q to quit: \n")
    if (userInput != 'q'):
        sum = sum+int(userInput)
        print(f"Order total so far: {sum}")

    else:

        print(f"Your Bill Total is {sum}. Thanks for shopping with us")
        break

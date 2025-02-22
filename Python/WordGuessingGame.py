import random
# username = input("Please Type Your Name")
# print("Hello", username)

print(f"Hello!, {input('please inter your name')}")
word = ["apple", "banana","mango","orange","kivi","drago","gorrilla","car","bik","bikeee","you","me","us","no","yes","why","for","do","this","dance","mr"]

selected = random.choice(word)
print(selected)

print("Guess the words to start  Game")
guessed_word =''
Traials = 12

while Traials>0:
    Failed =0

    for char in selected:
        if char in guessed_word:
            print("_",char)
        else:
            print("_")
            Failed += 1
    if Failed ==0:
        print("you won")
        print("the ord is ",selected)
        break
        

    print()
    guess=input("guess a word")
    guessed_word += guess

    if guess not in word:
        Traials -=1
        print(f"wrong, you have {Traials} ")

        if Traials ==0:
            print(f"you loose, the word was{selected}")






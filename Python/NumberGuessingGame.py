import random

print("Hey there, welcome ot no geussing game, You have 7 Opporunities... \nso guess it wisely. Lets STart the Game")

RandomNumber = random.randint(1, 100)
Guesses = 7
Guess_Counter = 0

while Guess_Counter < Guesses:
    Guess_Counter+=1
    my_guess = int(input("Your Turn : Plese give your guess"))
    if my_guess == RandomNumber:
        print(f"your guess{my_guess}, computer guess {RandomNumber} is matching, you won in {Guess_Counter} attempt")
        break
    
    elif (Guess_Counter>=Guesses and my_guess!=RandomNumber):
        print(f"you have exhausted your trials, your tril is {Guess_Counter}, play game again to win")
    elif (my_guess>RandomNumber):
        print(f"your guess is higher")
    elif (my_guess<RandomNumber):
        print(f"your guess is lower")





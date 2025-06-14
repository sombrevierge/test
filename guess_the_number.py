import random


def main():
    number = random.randint(1, 100)
    attempts = 0
    print("Guess the number between 1 and 100")
    while True:
        try:
            guess = int(input("Your guess: "))
        except ValueError:
            print("Please enter a valid integer.")
            continue
        attempts += 1
        if guess < number:
            print("Too low!")
        elif guess > number:
            print("Too high!")
        else:
            print(f"Correct! You guessed it in {attempts} tries.")
            break


if __name__ == "__main__":
    main()

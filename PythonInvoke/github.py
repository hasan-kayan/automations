import random
import subprocess

# Read the commit.txt file and retrieve existing words
def read_commit_file():
    try:
        with open("commit.txt", "r") as file:
            words = file.read().split()
            return words
    except FileNotFoundError:
        return []

# Write a random word to the commit.txt file
def write_random_word(words):
    with open("commit.txt", "a") as file:
        random_word = random.choice(["apple", "banana", "cherry", "date", "elderberry"])
        file.write(random_word + "\n")
        words.append(random_word)
        print(f"Added word: {random_word}")

# Read GitHub account information from info.txt
def read_github_info():
    try:
        with open("info.txt", "r") as file:
            info = file.read().splitlines()
            if len(info) == 2:
                return info[0], info[1]
            else:
                print("Invalid format in info.txt. Please provide username and repository name on separate lines.")
                return None, None
    except FileNotFoundError:
        print("info.txt not found.")
        return None, None

# Commit the changes to the specified GitHub repository
def commit_to_github(username, repository):
    subprocess.call(["git", "add", "commit.txt"])
    subprocess.call(["git", "commit", "-m", "Added random word"])
    subprocess.call(["git", "push", "-u", f"https://github.com/{username}/{repository}.git", "main"])

# Main script execution
def main():
    words = read_commit_file()
    write_random_word(words)
    username, repository = read_github_info()
    
    if username and repository:
        commit_to_github(username, repository)

if __name__ == "__main__":
    main()

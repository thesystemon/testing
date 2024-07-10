import subprocess
import sys

def main():
    print("Choose an option:")
    print("1. Run basic.py")
    print("2. Run extended-module.py")
    
    choice = input("Enter your choice (1 or 2): ")

    if choice == '1':
        print("Running basic.py...")
        subprocess.run([sys.executable, 'basic-module.py'])
    elif choice == '2':
        print("Running extended-module.py...")
        subprocess.run([sys.executable, 'extended-module.py'])
    else:
        print("Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()

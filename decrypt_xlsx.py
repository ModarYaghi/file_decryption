import itertools
import msoffcrypto
import pandas as pd
import io
from multiprocessing import Pool

def try_password(password):
    filename = "fc_bbi_rn_id.xlsx"
    try:
        file = msoffcrypto.OfficeFile(open(filename, "rb"))
        file.load_key(password=password) # Use password to decrypt
        decrypted = io.BytesIO()
        file.decrypt(decrypted)

        # Now you can use pandas to read the decrypted content:
        decrypted.seek(0)
        df = pd.read_excel(decrypted)
        print("File decrypted with password:", password)
        return password
    except Exception as e:
        return None

if __name__ == "__main__":
    chars = input("Enter characters: ")
    passwords = ["".join(combination) for length in range(1, len(chars) + 1) for combination in itertools.product(chars, repeat=length)]

    with Pool() as p:
        result = p.map(try_password, passwords)
        result = [r for r in result if r is not None]
        print("Passwords found:", result)

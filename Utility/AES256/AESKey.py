
def AES1():
    from Crypto.Cipher import AES
    obj = AES.new('This is a key123'.encode("utf8"), AES.MODE_CBC, 'This is an IV456'.encode("utf8"))

    message = "The answer is no".encode()
    print(obj.encrypt(message))

def AES2():
    import hashlib, binascii
    dk = hashlib.pbkdf2_hmac('sha256',b'password',b'salt',100000)
    ba = binascii.hexlify(dk)
    print(ba)

def AES3():
    import pyaes
    key = "This_key_for_demo_purposes_only!".encode()
    iv = "InitializationVe".encode()

    aes = pyaes.AESModeOfOperationCBC(key, iv = iv)
    plaintext = "TextMustBe16Byte".encode()
    ciphertext = aes.encrypt(plaintext)
    print(repr(ciphertext))
    aes = pyaes.AESModeOfOperationCBC(key, iv = iv)
    decrypted = aes.decrypt(ciphertext)

    print(decrypted == plaintext,plaintext,ciphertext)

    with open("KEY.bin",'r') as file:
        line = file.readline()

        k = [b for b in line]
            
    print(len(k))
            

AES3()

from file_finder import *
from file_miner import *

if __name__ == '__main__':
    token,filepath,token_list = file_finder()
    file_miner(token,filepath,token_list)
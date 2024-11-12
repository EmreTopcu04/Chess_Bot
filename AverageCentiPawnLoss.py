import chess
from stockfish import Stockfish
from openpyxl import Workbook
import random


def isOurGameOver():
    return board.is_checkmate() or board.is_stalemate() or board.is_insufficient_material() or board.is_seventyfive_moves() or board.is_fivefold_repetition() or board.can_claim_draw() or board.is_fifty_moves() or board.can_claim_threefold_repetition() or board.is_repetition(3)   

#Change The Path

stockfish = Stockfish("../Discrete Mathematics II/stockfish/stockfish-windows-x86-64-avx2.exe")
stockfish.set_depth(5)
stockfish.set_skill_level(20)
stockfish.update_engine_parameters({
    "Debug Log File": "",
    "Contempt": 0,
    "Min Split Depth": 0,
    "Threads": 4, # More threads will make the engine stronger, but should be kept at less than the number of logical processors on your computer.
    "Ponder": "false",
    "Hash": 2048, # Default size is 16 MB. It's recommended that you increase this value, but keep it as some power of 2. E.g., if you're fine using 2 GB of RAM, set Hash to 2048 (11th power of 2).
    "MultiPV": 1,
    "Skill Level": 20,
    "Move Overhead": 10,
    "Minimum Thinking Time": 20,
    "Slow Mover": 100,
    "UCI_Chess960": "false",
    "UCI_LimitStrength": "false",
    "UCI_Elo": 1350
})


wb = Workbook()
ws_game = wb.active
ws_game.title = "Chess Games"
ws_turn = wb.create_sheet("Game And Turn")

 # Headers for game data table
ws_game.cell(row=1, column=1).value = "Game Number"
ws_game.cell(row=1, column=2).value = "Ply Number"
ws_game.cell(row=1, column=3).value = "Who Won"

ws_turn.cell(row=1, column=1).value = "Number"

# Game counter and turn counter
game_number = 1
row = 2
colForTurn = 2
while game_number < 10000:

    board = chess.Board()
    rowForTurn = 1
    ws_turn.cell(row=rowForTurn, column=colForTurn).value = "Game " +str(game_number)
    rowForTurn += 1

    while board.is_game_over() == False or isOurGameOver == True:
        stockfish.set_fen_position(board.fen())
        list = stockfish.get_top_moves()
        print(list)       
        try:
            if len(list) > 1:
                print(board)
                print("\n")
                
                if (board.turn == chess.WHITE):
                    print(list)
                    if (list[0]["Centipawn"] == None):
                            counter = 0
                            for i in range(len(list)):
                                if list[i]["Mate"] != None:
                                    counter += 1
                            if counter != 0:
                                list = list[:counter]

                    elif (list[0]["Centipawn"] < 0):
                        #Check if there are mates
                        counter = 0
                        for i in range(len(list)):
                            if list[i]["Mate"] != None:
                                counter += 1
                        if counter != 0:
                            list = list[:counter]
                            
                        #If there are no mates take the best ones
                        else:
                            threshHoldVal = list[0]["Centipawn"] + int(list[0]["Centipawn"]/2.0)
                            while i < len(list):
                                    if len(list) == 1:
                                        break;
                                    if list[i]["Centipawn"] < threshHoldVal:
                                        list.pop(i)
                                        i = 0
                                    else:
                                        i += 1
                            
                    elif (list[0]["Centipawn"] >= 0):
                        #Check if there are mates
                        counter = 0
                        for i in range(len(list)):
                            if list[i]["Mate"] != None:
                                counter += 1
                        if counter != 0:
                            list = list[:counter]
                            
                        #If there are no mates take the best ones
                        else:
                            threshHoldVal = list[0]["Centipawn"] - int(list[0]["Centipawn"]/2.0)
                            while i < len(list):
                                    if len(list) == 1:
                                        break;
                                    if list[i]["Centipawn"] < threshHoldVal:
                                        list.pop(i)
                                        i = 0
                                    else:
                                        i += 1
                

                elif (board.turn == chess.BLACK):
                        print(list)
                        if (list[0]["Centipawn"] == None):
                            counter = 0
                            for i in range(len(list)):
                                if list[i]["Mate"] != None:
                                    counter += 1
                            if counter != 0:
                                list = list[:counter]
                        elif (list[0]["Centipawn"] < 0):
                            #Check if there are mates
                            counter = 0
                            for i in range(len(list)):
                                if list[i]["Mate"] != None:
                                    counter += 1
                            if counter != 0:
                                list = list[:counter]
                                
                            #If there are no mates take the best ones
                            else: 
                                threshHoldVal = list[0]["Centipawn"] - int(list[0]["Centipawn"]/2.0)
                                while i < len(list):
                                        if len(list) == 1:
                                            break;
                                        if list[i]["Centipawn"] > threshHoldVal:
                                            list.pop(i)
                                            i = 0
                                        else:
                                            i += 1

                        elif (list[0]["Centipawn"] >= 0):
                            #Check if there are mates
                            counter = 0
                            for i in range(len(list)):
                                if list[i]["Mate"] != None:
                                    counter += 1
                            if counter != 0:
                                list = list[:counter]
                                
                            #If there are no mates take the best ones
                            else:
                                threshHoldVal = list[0]["Centipawn"] + int(list[0]["Centipawn"]/2.0)
                                while i < len(list):
                                    if len(list) == 1:
                                        break;
                                    if list[i]["Centipawn"] > threshHoldVal:
                                        list.pop(i)
                                        i = 0
                                    else:
                                        i += 1
                if (len(list) == 1):
                    ws_turn.cell(row=rowForTurn, column=colForTurn).value = 1
                    board.push_san(list[0]["Move"])
                    rowForTurn += 1
                else:
                    ws_turn.cell(row=rowForTurn, column=colForTurn).value = len(list)
                    board.push_san(list[int(random.random()*len(list))]["Move"])
                    rowForTurn += 1
            else:
                ws_turn.cell(row=rowForTurn, column=colForTurn).value = 1
                board.push_san(list[0]["Move"])
                rowForTurn += 1
                print(board)
                print("\n")    
        except(TypeError):
                print("ERROR OCCURED CHECK FOR SIDE")
                
                
        
    ply_number = board.ply()
    ws_game.cell(row=row, column=1).value = "Game " +str(game_number)
    ws_game.cell(row=row, column=2).value = ply_number
    if board.outcome().result() == "1-0":
        ws_game.cell(row=row, column=3).value = "White"
        print("White Won")
    elif board.outcome().result() == "0-1":
        ws_game.cell(row=row, column=3).value = "Black"
        print("Black Won")

    else:
        ws_game.cell(row=row, column=3).value = "Draw"
        print("Draw")
    row += 1
    game_number += 1
    colForTurn += 1
    print(board)
    

# Save the workbook
wb.save("chess_games_with_turns.xlsx")

print("Excel file created successfully!")

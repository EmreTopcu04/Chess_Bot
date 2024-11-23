import chess
from stockfish import Stockfish
from openpyxl import Workbook
import random
import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import os

# Function to check if the game is over
def isOurGameOver(board):
    return (
        board.is_checkmate() or 
        board.is_stalemate() or 
        board.is_insufficient_material() or 
        board.is_seventyfive_moves() or 
        board.is_fivefold_repetition()
    )

# Initialize Stockfish with the path to the engine and engine parameters
stockfish = Stockfish("./stockfish-windows-x86-64-avx2.exe")
stockfish.set_depth(10)
stockfish.set_skill_level(20)
stockfish.update_engine_parameters({
    "Threads": 4,
    "Hash": 8192,
    "MultiPV": 3,
    "Skill Level": 20,
    "Ponder": "true"
})

# Define piece values for evaluation
PIECE_VALUES = {
    'P': 1, 'N': 3, 'B': 3, 'R': 5, 'Q': 9, 'K': 0,  # White pieces
    'p': -1, 'n': -3, 'b': -3, 'r': -5, 'q': -9, 'k': 0  # Black pieces
}

# Create a new Excel workbook for saving game results
wb = Workbook()
ws_game = wb.active
ws_game.title = "Chess Games"
ws_game.cell(row=1, column=1).value = "Game Number"
ws_game.cell(row=1, column=2).value = "Ply Number"
ws_game.cell(row=1, column=3).value = "Who Won"

# Initialize counters
game_number = 1
row = 2

# Initialize the GUI window
root = tk.Tk()
root.title("Chess Game Status")
root.geometry("700x700")

# Status label to show who is winning
status_frame = tk.Frame(root, bg="black", bd=5, relief="ridge")
status_frame.pack(pady=20, padx=10, fill="x")

status_label = tk.Label(
    status_frame, 
    text="Status Of The Game!", 
    font=("Helvetica", 18, "bold"), 
    fg="white", 
    bg="green", 
    bd=5, 
    padx=20, 
    pady=10, 
    wraplength=500,  # Wrap text if it exceeds this width
    anchor="center"  # Center the text within the label
)
status_label.pack(fill="both", expand=True, pady=5)

canvas = tk.Canvas(root, width=480, height=480)
canvas.pack(padx=10, pady=10)

# Load piece images
piece_images = {}
piece_names = {
    'K': 'white-king', 'Q': 'white-queen', 'R': 'white-rook', 'B': 'white-bishop', 
    'N': 'white-knight', 'P': 'white-pawn', 'k': 'king', 'q': 'queen', 
    'r': 'rook', 'b': 'bishop', 'n': 'knight', 'p': 'pawn'
}
for piece, filename in piece_names.items():
    image = Image.open(os.path.join("icons", f"{filename}.png"))
    image = image.resize((30, 30))
    piece_images[piece] = ImageTk.PhotoImage(image)



# Function to evaluate the board state
def evaluate_board(board):
    score = 0
    for square in chess.SQUARES:
        piece = board.piece_at(square)
        if piece:
            score += PIECE_VALUES.get(piece.symbol(), 0)
    return score

def update_status_label(board):
    evaluation = evaluate_board(board)
    if evaluation > 0:
        status_label.config(text=f"White is winning by {evaluation} points")
    elif evaluation < 0:
        status_label.config(text=f"Black is winning by {abs(evaluation)} points")
    else:
        status_label.config(text="The game is evenly balanced")

# Function to update the chessboard on the GUI
def update_chessboard(board):
    canvas.delete("all")
    square_size = 60

    for row in range(8):
        for col in range(8):
            x1 = col * square_size
            y1 = row * square_size
            x2 = x1 + square_size
            y2 = y1 + square_size
            color = "#D2B48C" if (row + col) % 2 == 0 else "#8B4513"  
            canvas.create_rectangle(x1, y1, x2, y2, fill=color)

            square_index = (7 - row) * 8 + col
            piece = board.piece_at(square_index)
            
            if piece:
                piece_symbol = piece.symbol()

                if piece_symbol in piece_images:
                    canvas.create_image(
                        x1 + square_size / 2,
                        y1 + square_size / 2,
                        image=piece_images[piece_symbol]
                    )
                else:
                    canvas.create_text(
                        x1 + square_size / 2,
                        y1 + square_size / 2,
                        text=piece_symbol,
                        fill="red",
                        font=("Helvetica", 20)
                    )

# Function to start a new game
def start_game():
    global game_number, row

    while game_number <= 10:
        board = chess.Board()
        update_chessboard(board)
        update_status_label(board)
        root.update()
    
        # Play the game
        while not isOurGameOver(board):
            stockfish.set_fen_position(board.fen())
            move_list = stockfish.get_top_moves(3)
            move = random.choice(move_list)["Move"]
            board.push_san(move)

            update_chessboard(board)
            update_status_label(board)
            root.update()

        # Record the game result
        ply_number = board.ply()
        ws_game.cell(row=row, column=1).value = f"Game {game_number}"
        ws_game.cell(row=row, column=2).value = ply_number

        if board.outcome().result() == "1-0":
            ws_game.cell(row=row, column=3).value = "White"
        elif board.outcome().result() == "0-1":
            ws_game.cell(row=row, column=3).value = "Black"
        else:
            ws_game.cell(row=row, column=3).value = "Draw"

        row += 1
        game_number += 1

    # Save the workbook
    wb.save("chess_games_with_turns.xlsx")
    messagebox.showinfo("Game Over", "Chess games completed and saved!")

# Start button to initiate the game
start_button = tk.Button(
    root, text="Start Game", command=start_game, font=("Helvetica", 16, "bold"),
    fg="white", bg="green", relief="raised", bd=5, padx=20, pady=10,
    activebackground="darkgreen", activeforeground="white"
)
start_button.pack(pady=20)



# Run the tkinter event loop
root.mainloop()
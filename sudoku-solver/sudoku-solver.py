#!/usr/bin/env python3

# Sudoku board to solve
sudoku_board = [
    [" ", " ", " ", "7", " ", " ", "3", " ", "1"],
    ["3", " ", " ", "9", " ", " ", " ", " ", " "],
    [" ", "4", " ", "3", "1", " ", "2", " ", " "],
    [" ", "6", " ", "4", " ", " ", "5", " ", " "],
    [" ", " ", " ", " ", " ", " ", " ", " ", " "],
    [" ", " ", "1", " ", " ", "8", " ", "4", " "],
    [" ", " ", "6", " ", "2", "1", " ", "5", " "],
    [" ", " ", " ", " ", " ", "9", " ", " ", "8"],
    ["8", " ", "5", " ", " ", "4", " ", " ", " "]
]

# descriptive names for subboards
subboard_names = {
    (0, 0): "top left",
    (0, 3): "top middle",
    (0, 6): "top right",
    (3, 0): "middle left",
    (3, 3): "middle",
    (3, 6): "middle right",
    (6, 0): "bottom left",
    (6, 3): "bottom middle",
    (6, 6): "bottom right"
}

# get_next_cell returns the next cell in row order.
# If the current cell is the last cell, it returns (9, 0).
def get_next_cell(row, col):
    if col == 8:
        col = 0
        row += 1
    else:
        col += 1
    return (row, col)

# find_row_nums returns the set of possible numbers for the given row.
def find_row_nums(board, row):
    num_set = {"1", "2", "3", "4", "5", "6", "7", "8", "9"}
    for i in range(9):
        if board[row][i] != " ":
            num_set.remove(board[row][i])
    return num_set

# find_col_nums returns the set of possible numbers for the given column.
def find_col_nums(board, col):
    num_set = {"1", "2", "3", "4", "5", "6", "7", "8", "9"}
    for i in range(9):
        if board[i][col] != " ":
            num_set.remove(board[i][col])
    return num_set

# find_subboard_nums returns the set of possible numbers for the 3 x 3 subboard
# of the given row and column.
def find_subboard_nums(board, row, col):
    num_set = {"1", "2", "3", "4", "5", "6", "7", "8", "9"}
    row = row - row % 3
    col = col - col % 3
    for i in range(3):
        for j in range(3):
            if board[row + i][col + j] != " ":
                num_set.remove(board[row + i][col + j])
    return num_set

# check_board returns whether the given board is valid.
# If the board is invalid, it prints an error message.
def check_board(board):
    for i in range(9):
        try:
            find_row_nums(board, i)
        except KeyError:
            print("The board is invalid. Invalid numbers in row ", i + 1, \
                  ".", sep="")
            return False
        
        try:
            find_col_nums(board, i)
        except KeyError:
            print("The board is invalid. Invalid numbers in column ", i + 1, \
                  ".", sep="")
            return False
    
    for i in range(0, 9, 3):
        for j in range(0, 9, 3):
            try:
                find_subboard_nums(board, i, j)
            except KeyError:
                print("The board is invalid. Duplicate numbers in", \
                      subboard_names[(i, j)], "subboard.")
                return False
    return True

# print_board prints the given board.
def print_board(board):
    print("-" * 19)
    for i in range(9):
        print("|", end="")
        for j in range(9):
            if j % 3 == 2:
                print(board[i][j], end="|")
            else:
                print(board[i][j], end=" ")
        print()
        if i % 3 == 2:
            print("-" * 19)

# solve_sudoku returns whether there is a solution.
# If there is a solution, the board contains the solution.
def solve_sudoku(board):
    row = 0
    col = 0
    while board[row][col] != " ":
        row, col = get_next_cell(row, col)
        if row == 9:
            return True
    
    row_nums = find_row_nums(board, row)
    col_nums = find_col_nums(board, col)
    subboard_nums = find_subboard_nums(board, row, col)
    possible_nums = row_nums & col_nums & subboard_nums
    
    for num in possible_nums:
        board[row][col] = num
        if solve_sudoku(board):
            return True
    board[row][col] = " "
    return False

# main is called when executed as a script.
def main():
    if check_board(sudoku_board):
        if solve_sudoku(sudoku_board):
            print("The solution is:")
        else:
            print("There is no solution for this board:")
    print_board(sudoku_board)

if __name__ == "__main__":
    main()

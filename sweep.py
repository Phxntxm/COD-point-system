from src.games import games

sweep = {}

# Go through each game, see how much you could get with a full sweep of all ILs and categories
for game in games:
    points = 0
    for category in game._main_game_leaderboards:
        points += 10000 * (category.percentage / 100)
    for category in game._il_leaderboards:
        points += 10000 * (category.percentage / 100)

    sweep[game.game] = points

# Sort the dictionary by points
sorted_points = sorted(sweep.items(), key=lambda x: x[1], reverse=True)

# Print the results
for game in sorted_points:
    print(f"{game[0]}: {int(game[1])}")

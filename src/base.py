import math
import time
import warnings
from dataclasses import dataclass
from typing import Dict, List

import numpy as np
import requests
from openpyxl.worksheet.worksheet import Worksheet

BASE_URL = "https://www.speedrun.com/api/v1"


@dataclass
class Run:
    runner: str | None
    name: str
    place: int
    time: float
    points: int
    main: bool
    game: str


@dataclass
class Player:
    id: str
    name: str
    runs: List[Run]

    @property
    def total_points(self):
        return sum([run.points for run in self.runs])

    @property
    def main_points(self):
        return sum([run.points for run in self.runs if run.main])

    @property
    def il_points(self):
        return sum([run.points for run in self.runs if not run.main])


@dataclass
class MainGameLeaderboard:
    id: str
    name: str
    percentage: float
    params: Dict[str, str] | None = None


@dataclass
class ILLeaderboard:
    id: str
    name: str
    main_game: bool = False
    params: Dict[str, str] | None = None


# This one's the weird one, when some games have what should be
#  an IL in the main game section (CoD4 MHC, FNG, etc.)
@dataclass
class MainGameILLeaderboard:
    id: str
    name: str
    category: str


@dataclass
class LevelSectionLeaderboard:
    levels: List[ILLeaderboard]
    category: str
    percentage: float = 100
    params: Dict[str, str] | None = None
    name: str = "ILs"
    full_game_section: bool = False


players: Dict[str, Player] = {}


class CODBase:
    _deviation_adjustment = 1.22
    _points_curvature = 1.48
    _weight = 5.5
    _rank_scaling = 0.38
    _rank_decay = 0.82
    _baseline_weight = 0.4

    cap = 10000
    game: str
    _levels = []

    _main_game_leaderboards: List[MainGameLeaderboard] = []
    _il_leaderboards: List[LevelSectionLeaderboard] = []

    @property
    def levels(self):
        if not self._levels:
            url = f"{BASE_URL}/games/{self.game}/levels"
            data = self._request(url)["data"]
            self._levels = [level["id"] for level in data]

        return self._levels

    @staticmethod
    def _request(url: str, params: Dict | None = None):
        response = requests.get(url, params=params)

        if response.status_code > 299:
            if response.status_code == 420:
                print("\nRate limit hit, waiting 60 seconds...")
                time.sleep(60)
                return CODBase._request(url, params)
            # This one's a bit weird, they'll randomly return 503 for
            #  "deploying an update" but the endpoint works a moment later.
            #  I assume it's an incorrect message, and it's just server issues
            #  at the time, so simply wait a second and try again
            elif response.status_code == 503:
                print("\nServer error, waiting 10 seconds...")
                time.sleep(10)
                return CODBase._request(url, params)
            else:
                print(f"\nUnknown error requesting {url}:\n\n{response.text}")
                exit(1)

        return response.json()

    def _consolidate_players(self, data: Dict):
        # The way SRC handles embed is kinda dumb, because it provides the embed on a DIFFERENT
        #  resource, when it should... you know... embed it in the resource...
        #  they even call the keyword "embed" lmao
        _players = data["data"]["players"]["data"]

        for player in _players:
            if player["rel"] == "guest":
                continue

            if player["id"] not in players:
                players[player["id"]] = Player(
                    id=player["id"], runs=[], name=player["names"]["international"]
                )

    def _request_leaderboard(
        self, game: str, leaderboard: MainGameLeaderboard | ILLeaderboard
    ):
        url = f"{BASE_URL}/leaderboards/{game}/category/{leaderboard.id}"

        params = leaderboard.params

        if params is None:
            params = {}

        params["embed"] = "players"
        data = self._request(url, params=params)
        self._consolidate_players(data)
        return data

    def _request_level(
        self, game: str, level: ILLeaderboard, leaderboard: LevelSectionLeaderboard
    ):
        url = f"{BASE_URL}/leaderboards/{game}/level/{level.id}/{leaderboard.category}"

        params = leaderboard.params

        if params is None:
            params = {}

        _level_params = level.params

        if _level_params is None:
            _level_params = {}

        params.update(_level_params)

        params["embed"] = "players"
        data = self._request(url, params=params)
        self._consolidate_players(data)
        return data

    def _calculate_points(
        self,
        runs: List[Run],
        leaderboard: MainGameLeaderboard | LevelSectionLeaderboard,
    ):
        # If there are no runs, obviously no points
        if len(runs) == 0:
            return

        # First, set the cap to the percentage of the total cap
        cap = self.cap * (leaderboard.percentage / 100)
        # Next dilute if this is for ILs
        if isinstance(leaderboard, LevelSectionLeaderboard):
            cap = cap / len(leaderboard.levels)

        # If there's only one run, or all the times are the same (two that are tied does happen)
        #  then give them 75% of the cap
        if len(runs) == 1 or len(set([run.time for run in runs])) == 1:
            for run in runs:
                if run.runner is None:
                    return

                players[run.runner].runs.append(run)
                run.points = math.floor(cap * 0.75)
            return

        wr = runs[0].time

        times = np.array([run.time for run in runs[:100]])
        std = np.std(times)
        V = std / wr

        for run in runs:
            M = self._deviation_adjustment * (run.time - wr) / std
            Q = self._rank_scaling * (run.place - 1) ** self._rank_decay
            w = 1 - (V + (1 - self._baseline_weight) ** (-1 / self._weight)) ** (
                0 - self._weight
            )
            D = w * Q + (1 - w) * M
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                run.points = math.floor(
                    cap
                    * (1 - (0.5 + 0.5 * math.erf(np.log(D) / self._points_curvature)))
                )
            if run.runner is None:
                continue

            players[run.runner].runs.append(run)

    def _handle_runs(
        self,
        runs: Dict,
        leaderboard: MainGameLeaderboard | LevelSectionLeaderboard,
        *,
        level: ILLeaderboard | None = None,
    ):
        if isinstance(leaderboard, MainGameLeaderboard):
            name = leaderboard.name
        elif level is not None:
            name = f"{leaderboard.name} - {level.name}"
        else:
            name = leaderboard.name

        _runs: List[Run] = []

        for run in runs:
            # Don't get unverified runs
            if run["run"]["status"]["status"] != "verified":
                continue
            # Currently glitched, and gives obsolete runs, who have a place of 0
            if run["place"] == 0:
                continue

            runner = run["run"]["players"][0]
            runner_id = runner["id"] if runner["rel"] == "user" else None

            _runs.append(
                Run(
                    runner=runner_id,
                    place=run["place"],
                    time=run["run"]["times"]["primary_t"],
                    main=isinstance(leaderboard, MainGameLeaderboard),
                    points=0,
                    name=name,
                    game=self.game,
                )
            )

        self._calculate_points(
            _runs,
            leaderboard,
        )

    def _calculate_main_runs(self):
        count = len(self._main_game_leaderboards)

        for i, leaderboard in enumerate(self._main_game_leaderboards):
            print(f"\t Handling main runs. [{i + 1}/{count}]                ", end="\r")
            runs = self._request_leaderboard(self.game, leaderboard)["data"]["runs"]

            self._handle_runs(runs, leaderboard)

        print()

    def _calculate_ils(self):
        for leaderboard in self._il_leaderboards:
            count = len(leaderboard.levels)

            for i, level in enumerate(leaderboard.levels):
                print(
                    f"\t Handling ILs. [{i + 1}/{count}] {leaderboard.name} - {level.name}                ",
                    end="\r",
                )

                if level.main_game:
                    runs = self._request_leaderboard(self.game, level)["data"]["runs"]
                else:
                    runs = self._request_level(self.game, level, leaderboard)["data"][
                        "runs"
                    ]

                self._handle_runs(runs, leaderboard, level=level)

            print()

    @staticmethod
    def get_categories(game: str):
        # Just a convenience method to list all the category id and names
        url = f"{BASE_URL}/games/{game}/categories"
        data = CODBase._request(url)["data"]
        for category in data:
            print(f"{category['id']}: {category['name']} ({category['type']})")

    @staticmethod
    def get_variables(game: str):
        # Just a convenience method to list all the variable id and names
        url = f"{BASE_URL}/games/{game}/variables"
        data = CODBase._request(url)["data"]
        for variable in data:
            if variable["is-subcategory"]:
                print(f"{variable['id']}: {variable['name']}")
                for _id, _name in variable["values"]["choices"].items():
                    print(f"\t{_id} ({_name})")

    @staticmethod
    def get_levels(game: str):
        # Just a convenience method to list all the level id and names
        url = f"{BASE_URL}/games/{game}/levels"
        data = CODBase._request(url)["data"]
        print("_il_leaderboards = [")
        print("\tLevelSectionLeaderboard(")
        print("\t\tlevels=[")
        for level in data:
            print(
                f"\t\t\tILLeaderboard(id=\"{level['id']}\", name=\"{level['name']}\"),"
            )
        print("\t\t],")
        print("\t),")
        print("]")

    def calculate(self):
        print(f"Calculating {self.game} points...")
        self._calculate_main_runs()
        self._calculate_ils()

from bs4 import BeautifulSoup
import openpyxl
import requests


def soup_session(link):
    """BeautifulSoup session"""
    session = requests.Session().get(link)
    soup = BeautifulSoup(session.content, 'html.parser')
    return soup


def get_deck_list(session):
    max_decks = 3
    deck_count = 1
    for deck in session.find_all('table', attrs={'class': 'listing listing-deckhash b-table b-table-a'}):
        for link in list(deck.find_all('a', href=True)):
            if "/top-decks/" in link['href']:
                get_deck_info(hearthpwn_link + link['href'].split("/top-decks")[1])
                if deck_count == max_decks:
                    return
                else:
                    deck_count += 1


def get_deck_info(deck_link):
    deck_info_dictionary = {}
    deck_link_session = soup_session(deck_link)

    # Returns deck class used (found in card table on right)
    for deck in deck_link_session.find_all('section', attrs={'class': 't-deck-details-card-list class-listing'}):
        deck_class = deck.find('h4').text.split(' ')[0]
        break

    for deck in deck_link_session.find_all('table', attrs={'class': 'listing listing-cards-tabular b-table b-table-a'}):
        for card in deck.find_all('a', href=True):
            try:
                deck_info_dictionary[card.text.strip()] = (card["data-count"], card["data-dust"])
            except KeyError:
                continue

    calculate_mana_cost_per_deck(deck_class, deck_info_dictionary)


def calculate_mana_cost_per_deck(deck_class, deck_collection):

    deck_collection_filepath = r"C:\Users\Riccardo\Desktop\Python Scripts\HearthStone Deck Maker\hearhstone_card_collection.xlsx"

    my_deck_collection = openpyxl.load_workbook(deck_collection_filepath)


if __name__ == "__main__":
    hearthpwn_link = "https://www.hearthpwn.com/top-decks"
    top_decks_session = soup_session(hearthpwn_link)
    get_deck_list(top_decks_session)

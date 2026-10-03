# Ethereum & Smart Contracts — Team 1 Presentation

This repository contains a 20-slide PowerPoint deck generator for an Ethereum, smart contracts, and DApps presentation with a dark blockchain-inspired visual style.

## Files

- `generate_presentation.py` — creates the deck as `Ethereum_Smart_Contracts_Team_1_Deck.pptx`
- `requirements.txt` — dependencies for the deck generator

## Generate the PowerPoint

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python generate_presentation.py
```

This will generate:

```text
Ethereum_Smart_Contracts_Team_1_Deck.pptx
```

## Deck structure

The generated presentation includes 20 slides covering:

1. Title
2. Agenda
3. Why blockchain matters
4. Ethereum overview
5. Ethereum architecture
6. Transactions and gas
7. Smart contracts introduction
8. Solidity basics
9. Remix IDE
10. Local development workflow
11. Ganache
12. DApp architecture
13. MetaMask and wallets
14. DeFi use case
15. NFT use case
16. Supply chain use case
17. Security and vulnerabilities
18. Scalability and future layers
19. Roadmap and opportunities
20. Closing / Q&A

## Design direction

- Dark technology palette
- Neon blockchain-inspired accents
- Clean card layouts and data blocks
- Web3 / DApp visual language
- Suitable for professional presentations and academic reports

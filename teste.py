import bitrix

card = bitrix.deal_get("32680")
print(type(card))
for key in card.keys():
    print(key)

print(card)
#236
from app import load_parameters

params = load_parameters()
print('Toplam parametre sayısı:', len(params))

kk_params = [p for p in params if p.get('KK')]
print('KK değeri olan parametre sayısı:', len(kk_params))

print('\nKK parametreleri:')
for p in kk_params[:10]:
    print(f"- {p.get('Parametre Adı')}: KK={p.get('KK')}, -3S={p.get('-3S')}, -2S={p.get('-2S')}, +2S={p.get('+2S')}, +3S={p.get('+3S')}")


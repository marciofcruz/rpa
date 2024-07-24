def main():
  numero1 = None
  numero2 = None
  numero3 = None
  
  print('Numero A:')
  numero1 = int(input())

  print('Numero B:')
  numero2 = int(input())
  
  print('Numero C:')
  numero3 = int(input())
  
  if numero1==numero2 and numero2==numero3:
    print('Todos os números são iguais')
  else:
    print(f'O maior número é {max(numero1, numero2, numero3)}')
  
  

if __name__ == "__main__":
  main()

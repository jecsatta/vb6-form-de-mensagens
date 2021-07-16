## VB6 Form de mensagens
![image](https://user-images.githubusercontent.com/60496134/125962549-2decf1c0-fc6f-4ed8-96e5-acbdcc316a30.png)
```` 
    With New FormMensagem
        .Titulo "Titulo"
        .Altura 3000 'optional
        .Largura 5000 'optional
        .Mensagem "Escolha uma Opcao"
        .AdicionaBotao "Teste"
        .AdicionaBotao "Teste2"
        a = .Mostra
    End With
````

Inline
````    
Set s = New FormMensagem
    a = s.Titulo("Titulo") _
        .Mensagem("Escolha uma Opcao") _
        .AdicionaBotao("Teste", 500) _
        .AdicionaBotao("Teste2", 3000) _
        .Mostra()
    Set s = Nothing
````

Attribute VB_Name = "modLanguage"
Public Sub Language()

    Select Case tmpCurLanguage
        
        ' Portugu�s
        Case 0
        
            ' Janlea de login
            TextUILoginUsername = "Usu�rio"
            TextUILoginPassword = "Senha"
            TextUILoginServerList = "Servidor"
            TextUILoginCheckBox = "Lembrar-me da senha?"
            TextUILoginEntryButton = "Entrar no PokeReborn"
            TextUILoginInvalidUsername = "Usu�rio inv�lido!"
            TextUILoginInvalidPassword = "Senha inv�lida!"
            
            ' Janela de registro
            TextUIRegisterUsername = "Usu�rio"
            TextUIRegisterPassword = "Senha"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar cadastro"
            TextUIRegisterCheckBox = "Mostrar a senha?"
            TextUIRegisterUsernameLenght = "Seu nome de usu�rio deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, n�meros e _ s�o permitidos."
            TextUIRegisterPasswordLenght = "Sua senha deve estar entre " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & "  caracteres e somente letras, n�meros e _ s�o permitidos."
            TextUIRegisterPasswordMatch = "A senhas n�o s�o iguais."
            TExtUIRegisterInvalidEmail = "Email inv�lido"
            
            ' Janela de cria��o de personagem
            TextUICreateCharacterCreateButton = "Finalizar personagem"
            TextUICreateCharacterUsername = "Nome"
            TextUICreateCharacterUsernameLenght = "Seu nome de personagem deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, n�meros e _ s�o permitidos"
            
            ' Mensagem universal
            TextUIWait = "Espere alguns segundos antes de tentar novamente."
            
            ' Footer
            TextUIFooterCreateAccount = "Criar uma nova conta"
            If CreditVisible Then
                TextUIFooterCredits = "Fechar"
            Else
                TextUIFooterCredits = "Cr�ditos"
            End If
            TextUIFooterDeveloper = "� PokeReborn - Todos os direitos reservados 2023."
            TextUIFooterChangePassword = "Mudar a senha"
            
            ' Menu global
            TextUIGlobalMenuReturn = "Retornar"
            TextUIGlobalMenuOptions = "Op��es"
            TextUIGlobalMenuReturnMenu = "Menu Principal"
            TextUIGlobalMenuExit = "Sair"
            
            ' Choice Window
            TextUIChoiceYes = "Sim"
            TextUIChoiceNo = "N�o"
            TextUIChoiceExit = "Tem certeza de que deseja sair do jogo?"
            TextUIChoiceReturnMainMenu = "Tem certeza que deseja retornar ao menu principal"
            TextUIChoiceBuyInvSlot = "Voc� quer comprar este slot por $" & INV_SLOTS_PRICE & " Cash?"
            TextUIChoiceEvolve = "Voc� quer evoluir seu pok�mon?"
            TextUIChoiceRelease = "Voc� tem certeza que deseja liberar este Pok�mon?"
            TextUIChoiceBuySlot = "Voc� quer comprar este slot por $" & Amount & "?"
            TextUIChoiceFly = "Voc� quer voar at� o local desta ins�gnia?"
            TextUIChoiceDuel = " convidou voc� para um duelo."
            TextUIChoiceTrade = " quer trocar com voc�."
            TextUIChoiceParty = " convidou voc� para um grupo."
            TextUIChoiceSave = "Deseja salvar essas configura��es?"
            TextUIChoiceDeleteCharacter = "Tem certeza de que deseja excluir este personagem?"
            
            ' Input Window
            TextUIInputAmountHeader = "Digite a quantidade:"
            TextUIInputNewPasswordHeader = "Digite a nova senha:"
            TextUIInputConfirm = "Confirmar"
            TextUIInputCancel = "Cancelar"
            
            ' Menu de op��es
            TextUIOptionVideoButton = "V�deo"
            TextUIOptionSoundButton = "Sons"
            TextUIOptionGameButton = "Jogo"
            TextUIOptionControlButton = "Controle"
            TextUIOptionFullscreen = "Tela Cheia: "
            TextUIOptionMusic = "M�sica"
            TextUIOptionSound = "Sons"
            TextUIOptionPath = "Interface:"
            TextUIOptionsFps = "Mostra o fps"
            TextUIOptionsPing = "Mostra o ping"
            TextUIOptionsFast = "In�cio R�pido"
            TextUIOptionName = "Mostrar Nome"
            TextUIOptionPP = "Mostrar PP Bar ao atacar"
            TextUIOptionLanguage = "Tradu��o: "
            TextUIOptionUp = "Subir"
            TextUIOptionDown = "Abaixo"
            TextUIOptionLeft = "Esquerda"
            TextUIOptionRight = "Direita"
            TextUIOptionCheckMove = "Movimentos"
            TextUIOptionMoveSlot1 = "Movimento 01"
            TextUIOptionMoveSlot2 = "Movimento 02"
            TextUIOptionMoveSlot3 = "Movimento 03"
            TextUIOptionMoveSlot4 = "Movimento 04"
            TextUIOptionAttack = "Atacar"
            TextUIOptionPokeSlot1 = "Pok�mon 01"
            TextUIOptionPokeSlot2 = "Pok�mon 02"
            TextUIOptionPokeSlot3 = "Pok�mon 03"
            TextUIOptionPokeSlot4 = "Pok�mon 04"
            TextUIOptionPokeSlot5 = "Pok�mon 05"
            TextUIOptionPokeSlot6 = "Pok�mon 06"
            TextUIOptionHotbarSlot1 = "Hotbar 01"
            TextUIOptionHotbarSlot2 = "Hotbar 02"
            TextUIOptionHotbarSlot3 = "Hotbar 03"
            TextUIOptionHotbarSlot4 = "Hotbar 04"
            TextUIOptionHotbarSlot5 = "Hotbar 05"
            TextUIOptionInventory = "Invent�rio"
            TextUIOptionPokedex = "Pok�dex"
            TextUIOptionInteract = "Interagir"
            TextUIOptionConvoChoice1 = "Con. Escolha 1"
            TextUIOptionConvoChoice2 = "Con. Escolha 2"
            TextUIOptionConvoChoice3 = "Con. Escolha 3"
            TextUIOptionConvoChoice4 = "Con. Escolha 4"
            
            ' Janela de Sele��o de personagem
            TextUICharactersNew = "Novo Personagem"
            TextUICharactersNone = "Vazio"
            TextUICharactersUse = "Usar"
            TextUICharactersDelete = "Del"
            
            ' Chat
            TextEnterToChat = "Aperte ENTER para digitar no chat"
            
        ' Ingl�s
        Case 1
        
            ' Login Window
            TextUILoginUsername = "User"
            TextUILoginPassword = "Password"
            TextUILoginServerList = "Server List"
            TextUILoginCheckBox = "Remember me password?"
            TextUILoginEntryButton = "Log in to PokeReborn"
            TextUILoginInvalidUsername = "Invalid username!"
            TextUILoginInvalidPassword = "Invalid password!"
            
            ' Choice Window
            TextUIChoiceYes = "Yes"
            TextUIChoiceNo = "No"
            TextUIChoiceExit = "Are you sure you want to exit the game?"
            TextUIChoiceReturnMainMenu = "Are you sure you want to return to main menu?"
            TextUIChoiceBuyInvSlot = "Do you want to buy this slot for $" & INV_SLOTS_PRICE & " Cash?"
            TextUIChoiceEvolve = "Do you want to evolve your pokemon?"
            TextUIChoiceRelease = "Are you sure that you want to release this Pokemon?"
            TextUIChoiceBuySlot = "Do you want to buy this slot for $" & Amount & "?"
            TextUIChoiceFly = "Do you want to fly to this badge's location?"
            TextUIChoiceDuel = " invites you to a duel"
            TextUIChoiceTrade = " wants to trade with you"
            TextUIChoiceParty = " invite you to a party"
            TextUIChoiceSave = "Do you want to save this settings?"
            TextUIChoiceDeleteCharacter = "Are you sure that you want to delete this character?"
            
            ' Input Window
            TextUIInputAmountHeader = "Enter amount:"
            TextUIInputNewPasswordHeader = "Enter your new password:"
            TextUIInputConfirm = "Confirm"
            TextUIInputCancel = "Cancel"
            
            ' Register Window
            TextUIRegisterUsername = "Usu�rio"
            TextUIRegisterPassword = "Senha"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar cadastro"
            TextUIRegisterCheckBox = "Mostrar a senha?"
            TextUIRegisterUsernameLenght = "Your username must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers and _ allowed"
            TextUIRegisterPasswordLenght = "Your password must be between " & ((NAME_LENGTH - 1) \ 4) & " and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers _ allowed"
            TextUIRegisterPasswordMatch = "Password did not match"
            TExtUIRegisterInvalidEmail = "Invalid email"
            
            ' Create Character Window
            TextUICreateCharacterCreateButton = "Create Character"
            TextUICreateCharacterUsername = "Name"
            TextUICreateCharacterUsernameLenght = "Your character name must be between 3 and " & (NAME_LENGTH - 1) & " characters long, and only letters, numbers and _ allowed"
                        
            'Text Universal
            TextUIWait = "Wait a few seconds before trying again."
            
            ' Footer
            TextUIFooterCreateAccount = "Create an account"
            If CreditVisible Then
                TextUIFooterCredits = "Close"
            Else
                TextUIFooterCredits = "Credits"
            End If
            TextUIFooterDeveloper = "� PokeReborn - Todos os direitos reservados 2023."
            TextUIFooterChangePassword = "Change password"
            
        ' Espanhol
        Case 2
            
            ' Login Window
            TextUILoginUsername = "Usuario"
            TextUILoginPassword = "Contrase�a"
            TextUILoginServerList = "Servidor"
            TextUILoginCheckBox = "Olvid� mi contrase�a?"
            TextUILoginEntryButton = "Entrar a PokeReborn"
            TextUILoginInvalidUsername = "Usuario incorrecto!"
            TextUILoginInvalidPassword = "Contrase�a incorrecta!"
            
            ' Choice Window
            TextUIChoiceYes = "S�"
            TextUIChoiceNo = "No"
            TextUIChoiceExit = "Est�s segura de que quieres salir del juego?"
            TextUIChoiceReturnMainMenu = "Are you sure you want to return to main menu?"
            TextUIChoiceBuyInvSlot = "Do you want to buy this slot for $" & INV_SLOTS_PRICE & " Cash?"
            TextUIChoiceEvolve = "Do you want to evolve your pokemon?"
            TextUIChoiceRelease = "Are you sure that you want to release this Pokemon?"
            TextUIChoiceBuySlot = "Do you want to buy this slot for $" & Amount & "?"
            TextUIChoiceFly = "Do you want to fly to this badge's location?"
            TextUIChoiceDuel = " invites you to a duel"
            TextUIChoiceTrade = " wants to trade with you"
            TextUIChoiceParty = " invite you to a party"
            TextUIChoiceSave = "Do you want to save this settings?"
            TextUIChoiceDeleteCharacter = "Are you sure that you want to delete this character?"
            
            ' Input Window
            TextUIInputAmountHeader = "Ingrese la cantidad:"
            TextUIInputNewPasswordHeader = "Introduzca su nueva contrase�a:"
            TextUIInputConfirm = "Confirmar"
            TextUIInputCancel = "Cancelar"
            
            ' Register Window
            TextUIRegisterUsername = "Usuario"
            TextUIRegisterPassword = "Contrase�a"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar el registro"
            TextUIRegisterCheckBox = "Mostrar contrase�a?"
            TextUIRegisterUsernameLenght = "Tu nombre de usuario debe tener entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y n�meros."
            TextUIRegisterPasswordLenght = "Tu contrase�a debe tener m�nimo " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y n�meros."
            TextUIRegisterPasswordMatch = "Las contrase�as no coinciden"
            TExtUIRegisterInvalidEmail = "Email no v�lido"
            
            ' Create Character Window
            TextUICreateCharacterCreateButton = "Finalizar el car�cter"
            TextUICreateCharacterUsername = "Nombre"
            TextUICreateCharacterUsernameLenght = "El nombre de tu personaje debe estar entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y n�meros."
            
            
            'Text Universal
            TextUIWait = "Espera unos segundos antes de intentar de nuevo."
            
            ' Footer
            TextUIFooterCreateAccount = "Crear una cuenta"
            If CreditVisible Then
                TextUIFooterCredits = "Cerrar"
            Else
                TextUIFooterCredits = "Cr�ditos"
            End If
            TextUIFooterDeveloper = "� PokeReborn - Todos os direitos reservados 2023."
            TextUIFooterChangePassword = "Cambiar contrase�a"
    End Select

End Sub

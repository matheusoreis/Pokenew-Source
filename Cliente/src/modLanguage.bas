Attribute VB_Name = "modLanguage"
Public Sub Language()

    Select Case tmpCurLanguage
        
        ' Português
        Case 0
        
            ' Janlea de login
            TextUILoginUsername = "Usuário"
            TextUILoginPassword = "Senha"
            TextUILoginServerList = "Servidor"
            TextUILoginCheckBox = "Lembrar-me da senha?"
            TextUILoginEntryButton = "Entrar no PokeReborn"
            TextUILoginInvalidUsername = "Usuário inválido!"
            TextUILoginInvalidPassword = "Senha inválida!"
            
            ' Janela de registro
            TextUIRegisterUsername = "Usuário"
            TextUIRegisterPassword = "Senha"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar cadastro"
            TextUIRegisterCheckBox = "Mostrar a senha?"
            TextUIRegisterUsernameLenght = "Seu nome de usuário deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, números e _ são permitidos."
            TextUIRegisterPasswordLenght = "Sua senha deve estar entre " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & "  caracteres e somente letras, números e _ são permitidos."
            TextUIRegisterPasswordMatch = "A senhas não são iguais."
            TExtUIRegisterInvalidEmail = "Email inválido"
            
            ' Janela de criação de personagem
            TextUICreateCharacterCreateButton = "Finalizar personagem"
            TextUICreateCharacterUsername = "Nome"
            TextUICreateCharacterUsernameLenght = "Seu nome de personagem deve estar entre 3 e " & (NAME_LENGTH - 1) & " caracteres e somente letras, números e _ são permitidos"
            
            ' Mensagem universal
            TextUIWait = "Espere alguns segundos antes de tentar novamente."
            
            ' Footer
            TextUIFooterCreateAccount = "Criar uma nova conta"
            If CreditVisible Then
                TextUIFooterCredits = "Fechar"
            Else
                TextUIFooterCredits = "Créditos"
            End If
            TextUIFooterDeveloper = "© PokeReborn - Todos os direitos reservados 2023."
            TextUIFooterChangePassword = "Mudar a senha"
            
            ' Menu global
            TextUIGlobalMenuReturn = "Retornar"
            TextUIGlobalMenuOptions = "Opções"
            TextUIGlobalMenuReturnMenu = "Menu Principal"
            TextUIGlobalMenuExit = "Sair"
            
            ' Choice Window
            TextUIChoiceYes = "Sim"
            TextUIChoiceNo = "Não"
            TextUIChoiceExit = "Tem certeza de que deseja sair do jogo?"
            TextUIChoiceReturnMainMenu = "Tem certeza que deseja retornar ao menu principal"
            TextUIChoiceBuyInvSlot = "Você quer comprar este slot por $" & INV_SLOTS_PRICE & " Cash?"
            TextUIChoiceEvolve = "Você quer evoluir seu pokémon?"
            TextUIChoiceRelease = "Você tem certeza que deseja liberar este Pokémon?"
            TextUIChoiceBuySlot = "Você quer comprar este slot por $" & Amount & "?"
            TextUIChoiceFly = "Você quer voar até o local desta insígnia?"
            TextUIChoiceDuel = " convidou você para um duelo."
            TextUIChoiceTrade = " quer trocar com você."
            TextUIChoiceParty = " convidou você para um grupo."
            TextUIChoiceSave = "Deseja salvar essas configurações?"
            TextUIChoiceDeleteCharacter = "Tem certeza de que deseja excluir este personagem?"
            
            ' Input Window
            TextUIInputAmountHeader = "Digite a quantidade:"
            TextUIInputNewPasswordHeader = "Digite a nova senha:"
            TextUIInputConfirm = "Confirmar"
            TextUIInputCancel = "Cancelar"
            
            ' Menu de opções
            TextUIOptionVideoButton = "Vídeo"
            TextUIOptionSoundButton = "Sons"
            TextUIOptionGameButton = "Jogo"
            TextUIOptionControlButton = "Controle"
            TextUIOptionFullscreen = "Tela Cheia: "
            TextUIOptionMusic = "Música"
            TextUIOptionSound = "Sons"
            TextUIOptionPath = "Interface:"
            TextUIOptionsFps = "Mostra o fps"
            TextUIOptionsPing = "Mostra o ping"
            TextUIOptionsFast = "Início Rápido"
            TextUIOptionName = "Mostrar Nome"
            TextUIOptionPP = "Mostrar PP Bar ao atacar"
            TextUIOptionLanguage = "Tradução: "
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
            TextUIOptionPokeSlot1 = "Pokémon 01"
            TextUIOptionPokeSlot2 = "Pokémon 02"
            TextUIOptionPokeSlot3 = "Pokémon 03"
            TextUIOptionPokeSlot4 = "Pokémon 04"
            TextUIOptionPokeSlot5 = "Pokémon 05"
            TextUIOptionPokeSlot6 = "Pokémon 06"
            TextUIOptionHotbarSlot1 = "Hotbar 01"
            TextUIOptionHotbarSlot2 = "Hotbar 02"
            TextUIOptionHotbarSlot3 = "Hotbar 03"
            TextUIOptionHotbarSlot4 = "Hotbar 04"
            TextUIOptionHotbarSlot5 = "Hotbar 05"
            TextUIOptionInventory = "Inventário"
            TextUIOptionPokedex = "Pokédex"
            TextUIOptionInteract = "Interagir"
            TextUIOptionConvoChoice1 = "Con. Escolha 1"
            TextUIOptionConvoChoice2 = "Con. Escolha 2"
            TextUIOptionConvoChoice3 = "Con. Escolha 3"
            TextUIOptionConvoChoice4 = "Con. Escolha 4"
            
            ' Janela de Seleção de personagem
            TextUICharactersNew = "Novo Personagem"
            TextUICharactersNone = "Vazio"
            TextUICharactersUse = "Usar"
            TextUICharactersDelete = "Del"
            
            ' Chat
            TextEnterToChat = "Aperte ENTER para digitar no chat"
            
        ' Inglês
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
            TextUIRegisterUsername = "Usuário"
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
            TextUIFooterDeveloper = "© PokeReborn - Todos os direitos reservados 2023."
            TextUIFooterChangePassword = "Change password"
            
        ' Espanhol
        Case 2
            
            ' Login Window
            TextUILoginUsername = "Usuario"
            TextUILoginPassword = "Contraseña"
            TextUILoginServerList = "Servidor"
            TextUILoginCheckBox = "Olvidé mi contraseña?"
            TextUILoginEntryButton = "Entrar a PokeReborn"
            TextUILoginInvalidUsername = "Usuario incorrecto!"
            TextUILoginInvalidPassword = "Contraseña incorrecta!"
            
            ' Choice Window
            TextUIChoiceYes = "Sí"
            TextUIChoiceNo = "No"
            TextUIChoiceExit = "Estás segura de que quieres salir del juego?"
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
            TextUIInputNewPasswordHeader = "Introduzca su nueva contraseña:"
            TextUIInputConfirm = "Confirmar"
            TextUIInputCancel = "Cancelar"
            
            ' Register Window
            TextUIRegisterUsername = "Usuario"
            TextUIRegisterPassword = "Contraseña"
            TextUIRegisterEmail = "Email"
            TextUIRegisterConfirm = "Finalizar el registro"
            TextUIRegisterCheckBox = "Mostrar contraseña?"
            TextUIRegisterUsernameLenght = "Tu nombre de usuario debe tener entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y números."
            TextUIRegisterPasswordLenght = "Tu contraseña debe tener mínimo " & ((NAME_LENGTH - 1) / 4) & " and " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y números."
            TextUIRegisterPasswordMatch = "Las contraseñas no coinciden"
            TExtUIRegisterInvalidEmail = "Email no válido"
            
            ' Create Character Window
            TextUICreateCharacterCreateButton = "Finalizar el carácter"
            TextUICreateCharacterUsername = "Nombre"
            TextUICreateCharacterUsernameLenght = "El nombre de tu personaje debe estar entre 3 y " & (NAME_LENGTH - 1) & " caracteres de largo, solo se permiten letras y números."
            
            
            'Text Universal
            TextUIWait = "Espera unos segundos antes de intentar de nuevo."
            
            ' Footer
            TextUIFooterCreateAccount = "Crear una cuenta"
            If CreditVisible Then
                TextUIFooterCredits = "Cerrar"
            Else
                TextUIFooterCredits = "Créditos"
            End If
            TextUIFooterDeveloper = "© PokeReborn - Todos os direitos reservados 2023."
            TextUIFooterChangePassword = "Cambiar contraseña"
    End Select

End Sub

Attribute VB_Name = "StructureItem"
Public Item(1 To MAX_ITEM) As ItemModel

Private Type ItemGacha
    ItemNumber As Integer                                   ' Armazena o número do item
    ItemValue As Long                                       ' Armazena o valor do item
    ItemChance As Double                                    ' Armazena a chance do item
End Type

Private Type ItemCooldown
    Value As Long                                           ' Armazena o valor do cooldown (em tempo)
    Type As Byte                                            ' Armazena o tipo do cooldown (por exemplo, segundos, minutos)
End Type

Private Type ItemRestriction
    CanStack As Boolean                                     ' Indica se o item pode ser empilhado
    CanHold As Boolean                                      ' Indica se o item pode ser mantido no inventário
    IsConnected As Boolean                                  ' Indica se o item está conectado (fisicamente ou de outra forma)
    IsAdminItem As Boolean                                  ' Indica se o item pode ser usado apenas por administradores
End Type

Private Type ItemRequirementPokemon
    RequiredLevel As Byte                                   ' Armazena o nível necessário
    PrimaryType As Byte                                     ' Armazena o tipo primário necessário
    SecondaryType As Byte                                   ' Armazena o tipo secundário necessário
End Type

Private Type ItemRequirementPlayer
    RequiredLevel As Byte                                   ' Armazena o nível necessário ao jogador
    RequiredMaps(1 To MAX_MAPS_REQUIREMENTS) As Long        ' Armazena os mapas necessários
    RequiredBadge As Byte                                   ' Armazena a insígnia necessária
End Type

Private Type ItemPokeball
    CaptureChance As Byte                                   ' Armazena a chance de captura
    SpriteID As Long                                        ' Armazena o ID da sprite
    HasPerfectCapture As Boolean                            ' Indica se a Pokébola permite captura perfeita
End Type

Private Type ItemMedicine
    Type As Byte                                            ' Armazena o tipo da medicina
    Value As Long                                           ' Armazena o valor da medicina
    HasLeveledUp As Boolean                                 ' Indica se aumenta o nível
End Type

Private Type ItemProteins
    Type As Byte                                            ' Armazena o tipo da proteína
    Value As Long                                           ' Armazena o valor da proteína
End Type

Private Type ItemKey
    Type As Byte                                            ' Armazena o tipo da chave
    Sprite As Byte                                          ' Armazena o ID da sprite
    ExperienceBonusAmount As Byte                           ' Armazena o bônus de experiência
    MoneyBonusAmount As Byte                                ' Armazena o bônus de experiência
    IsShiftRunning As Boolean                               ' Indica se é possível correr com o Shift
End Type

Private Type ItemSkills
    Type As Long                                            ' Armazena o tipo da habilidade
    CanConsume As Boolean                                   ' Indica se pode ser consumida
End Type

Private Type ItemBracelet
    Type As Byte                                            ' Armazena o tipo do bracelete
    Value As Long                                           ' Armazena o valor do bracelete
End Type

Private Type ItemModel
    Name As String * NAME_LENGTH                            ' Armazena o nome do item
    SpriteID As Long                                        ' Armazena o ID da sprite do item
    Rarity As Byte                                          ' Armazena a raridade do item
    Category As Byte                                        ' Armazena a categoria do item
    ExecutionType As Byte                                   ' Armazena o tipo de execução do item
    Description As String * 255                             ' Armazena a descrição do item
    
    CooldownData As ItemCooldown
    RestrictionData As ItemRestriction
    PokemonRequirementData As ItemRequirementPokemon
    PlayerRequirementData As ItemRequirementPlayer
    PokeballData As ItemPokeball
    MedicineData As ItemMedicine
    ProteinsData As ItemProteins
    KeyData As ItemKey
    SkillsData As ItemSkills
    BraceletData As ItemBracelet
    GachaData(1 To MAX_MYSTERY_BOX) As ItemGacha
    
End Type

Public Enum KeyTypeEnum
    None = 0
    Sprite
    OpenBank
    OpenComputer
    
    Count
End Enum

Public Enum BadgeEnum
    None = 0
    Boulder
    Cascade
    Thunder
    Rainbow
    Soul
    Marsh
    Volcano
    Earth
    
    Count
End Enum

Public Enum RarityCategoryEnum
    None = 0
    Nada
    Uncommon
    Rare
    VeryRare
    Legendary
    
    Count
End Enum

Public Enum ActionEnum
    None = 0
    OpenBank
    OpenComputer
    
    Count
End Enum

Public Enum ItemCategoryEnum
    None = 0
    PokeBall
    Medicine
    Protein
    Key
    Skills
    Bracelet
    Gacha
    
    Count
End Enum


Attribute VB_Name = "modFish"
Option Explicit

Private mapPokemonIds(MAX_MAP) As Collection

Private Sub Initialize()
    Dim i As Long
    For i = 1 To MAX_MAP
        Set mapPokemonIds(i) = New Collection
    Next i
End Sub

Public Sub AddPokemonIdToMap(ByVal mapIndex As Long, ByVal pokemonId As Long)
    mapPokemonIds(mapIndex).Add pokemonId
End Sub

Public Sub SpawnPokemonIdInMap(ByVal Index As Long, ByVal mapIndex As Long)
    Dim pokemonIds As Collection
    Dim pokemonId As Variant
    Dim Rand As Long, CountChances As Long

    Set pokemonIds = mapPokemonIds(mapIndex)

    If pokemonIds.count > 0 Then
RandomizaNovamente:
        Rand = Random(1, CLng(pokemonIds.count))
        
        '//Limitar a quantidade de tentativas em 30, pra não ter problemas pra sair daqui.
        If CountChances > 30 Then Exit Sub
        
        If IsWithinSpawnTime(pokemonIds.Item(Rand), GameHour) Then
            Call SpawnMapPokemon(pokemonIds.Item(Rand), , , Index)
        Else
            GoTo RandomizaNovamente
            CountChances = CountChances + 1
        End If
        
    End If
End Sub

Public Sub AddPokemonsFishing()
    Dim i As Long
    Initialize

    For i = 1 To MAX_GAME_POKEMON
        If Spawn(i).PokeNum > 0 And Spawn(i).Fishing And Spawn(i).MapNum > 0 Then
            AddPokemonIdToMap Spawn(i).MapNum, i
        End If
    Next i
End Sub

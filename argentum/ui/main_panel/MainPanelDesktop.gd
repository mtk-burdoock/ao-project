extends Control

@onready var stats_bars = $PanelContainerBottom/VBoxContainer2/StatsBars
@onready var spells_container = $PanelContainerTop/MidlePanel/Spells
@onready var items_container = $PanelContainerTop/MidlePanel/Inventario/ItemsContainer

func initialize(player_data:PlayerData, protocol:GameProtocol) -> void:
	var stats = player_data.stats
	if player_data && protocol:
		stats_bars.initialize(player_data, protocol)
		items_container.initialize(player_data, protocol)
		#spells_container.intialize(player_data.stats, protocol)

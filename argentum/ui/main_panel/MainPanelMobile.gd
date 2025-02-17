extends Control

@onready var _spellContainer = find_child("SpellContainerMobile")
@onready var _inventoryContainer = find_child("InventoryContainerMobile")
@onready var _ntnSwitchPanel = find_child("BtnSwitchPanel")
@onready var stats_bars = $StatsPanel/StatsBars

const _spellTexture = preload("res://assets/graphics/531.png")
const _inventoryTexture = preload("res://assets/graphics/572.png")

func initialize(player_data:PlayerData, protocol:GameProtocol) -> void:
	var stats = player_data.stats
	if player_data && protocol:
		stats_bars.initialize(player_data, protocol)
		_inventoryContainer.initialize(player_data, protocol)
		_spellContainer.intialize(player_data.stats, protocol)

func _on_BtnSwitchPanel_pressed() -> void:
	if _spellContainer.visible:
		_inventoryContainer.visible = true
		_spellContainer.visible = false
		_ntnSwitchPanel.texture_normal = _spellTexture
	else:
		_inventoryContainer.visible = false
		_spellContainer.visible = true 
		_ntnSwitchPanel.texture_normal = _inventoryTexture

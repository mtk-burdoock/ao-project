extends CanvasLayer

@onready var spellContainer = find_child("SpellContainerGUI")
@onready var inventoryContainer = find_child("InventoryContainerGUI")

#@onready var main_panel = $VBoxContainer/MainPanel
#@onready var mobile_main_panel = $VBoxContainer/MainPanelMobile
#@onready var rich_text_label = $RichTextLabel

var _protocol:GameProtocol
var _player_data:PlayerData
var _main_panel:Control
var _rich_text_label:RichTextLabel

func initialize(player_data:PlayerData, protocol:GameProtocol, main_panel:Control, rich_text_label:RichTextLabel) -> void:	
	
	_protocol = protocol
	_player_data = player_data
	_main_panel = main_panel
	_rich_text_label = rich_text_label
	
	protocol.connect("parse_data", Callable(self, "_on_parse_data"))
	if !is_inside_tree():
		await self.ready
	_main_panel.initialize(_player_data, protocol)
	
func _on_parse_data(packet_id, data):
	match packet_id:
		GameProtocol.ServerPacketID.ConsoleMsg:
			_rich_text_label.text += data.message + "\n"
	
func _unhandled_input(event: InputEvent) -> void:
	if event.is_action_pressed("toggle_combat_mode"):
		_protocol.write_combat_mode_toggle()
	
	if event.is_action_pressed("attack"):
		if !_player_data.timers[PlayerData.TimersIndex.Arrows].check(false): return
		
		if !_player_data.timers[PlayerData.TimersIndex.CastSpell].check(false):
			if !_player_data.timers[PlayerData.TimersIndex.CastAttack].check(): return
		else:
			if !_player_data.timers[PlayerData.TimersIndex.Attack].check(): return
			
				
		_protocol.write_attack()
	
	if event.is_action_pressed("pickup"):
		_protocol.write_pick_up() 
	
	if event.is_action_pressed("hide"):
		_protocol.write_work(Global.eSkill.Ocultarse) 
	
	if event.is_action_pressed("use_object"):
		if inventoryContainer.selected_slot:
			if _player_data.timers[PlayerData.TimersIndex.UseItemWithU].check():
				_protocol.write_use_item(inventoryContainer.selected_slot.slot_index + 1)

	if event.is_action_pressed("exit_game"):
		_protocol.write_quit() 
		
	if event.is_action_pressed("request_refresh"):
		if _player_data.check_timer(PlayerData.TimersIndex.SendRPU): 
			_protocol.write_request_position_update()
		
func _on_TouchScreenButton_pressed() -> void:
	print("pressed")

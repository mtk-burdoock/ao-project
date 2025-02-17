extends MarginContainer 

@onready var grid_container = $GridContainer

var _protocol:GameProtocol
var _player_data:PlayerData
var slot_selected:ItemSlot

func initialize(player_data:PlayerData, protocol:GameProtocol): 
	_player_data = player_data
	_protocol = protocol
	
	player_data.inventory.connect("slot_changed", Callable(self, "_on_slot_changed"))

	var i = 0
	for slot in grid_container.get_children():
		slot = slot as ItemSlot
		slot.inventory_index = i
		slot.connect("item_selected", Callable(self, "_on_item_selected").bind(slot))
		slot.selectedLabel.visible = false
		i += 1
	
	
func _on_slot_changed(slot:int, old_content:ItemStack, new_context:ItemStack) -> void:
	for i in grid_container.get_children():
		if i is ItemSlot and i.inventory_index == slot:
			i.set_item(slot, new_context.item, new_context.quantity, new_context.equipped)
			break
			
func _on_item_selected(itemslot:ItemSlot) -> void:		
	var index = itemslot.inventory_index + 1
	var item = itemslot.item

	if item and item.type:
		if item.is_consumable():
			_protocol.write_use_item(index)
		elif item.is_equippable():
			_protocol.write_equip_item(index)
		

"""
CODIGO VIEJO:
const ITEM_SLOT_SCENE = preload("res://ui/inventory/ItemSlot.tscn")

@onready var itemGrid = find_child("ItemGrid")

var inventory:Inventory = null
var slot_selected:int = -1

func set_inventory(inventory:Inventory) -> void:
	self.inventory = inventory
	
	inventory.connect("slot_changed", Callable(self, "_on_inventory_changed"))
	
	for i in inventory.max_slots:
		var stack = inventory.get_item_stack(i)
		var slot = ITEM_SLOT_SCENE.instantiate()
		itemGrid.add_child(slot)
		
		slot.set_item(i, stack.item, stack.quantity, stack.equipped)
		#slot.connect("item_selected", self, "_on_item_slot_selected", [slot])

func _on_item_slot_selected(slot:ItemSlot):
	print(slot.inventory_index)

func _on_inventory_changed(index:int, old_content:ItemStack, new_content:ItemStack) -> void:
	for child in itemGrid.get_children():
		if not child is ItemSlot:
			continue
		
		if child.inventory_index == index:
			child.set_item(index, new_content.item, new_content.quantity, new_content.equipped)
			break
"""

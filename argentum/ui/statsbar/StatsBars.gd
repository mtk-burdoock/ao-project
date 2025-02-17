extends VBoxContainer

@onready var progress_bar_hp = $ProgressBarHP
@onready var progress_bar_mp = $ProgressBarMP
@onready var progress_bar_sta = $ProgressBarSTA
@onready var progress_bar_ham = $HBoxHealth/ProgressBarHAM
@onready var progress_bar_sed = $HBoxHealth/ProgressBarSED
@onready var gold_label = $HBoxGold/GoldLabel

func initialize(player_data:PlayerData, protocol:GameProtocol) -> void:	
	var stats = player_data.stats
	
	stats.connect("change_hp", Callable(self, "_on_change_hp"))
	stats.connect("change_mp", Callable(self, "_on_change_mp"))
	stats.connect("change_sta", Callable(self, "_on_change_sta"))
	stats.connect("change_ham", Callable(self, "_on_change_ham"))
	stats.connect("change_sed", Callable(self, "_on_change_sed"))
	stats.connect("change_gold", Callable(self, "_on_change_gold"))

func _on_change_hp(value:int, max_value:int) -> void:
	progress_bar_hp.value = value
	progress_bar_hp.max_value = max_value

func _on_change_mp(value:int, max_value:int) -> void:
	progress_bar_mp.value = value
	progress_bar_mp.max_value = max_value

func _on_change_sta(value:int, max_value:int) -> void:
	progress_bar_sta.value = value
	progress_bar_sta.max_value = max_value

func _on_change_ham(value:int, max_value:int) -> void:
	progress_bar_ham.value = value
	progress_bar_ham.max_value = max_value

func _on_change_sed(value:int, max_value:int) -> void:
	progress_bar_sed.value = value
	progress_bar_sed.max_value = max_value

func _on_change_gold(value:int) -> void:
	gold_label.text = str(value)

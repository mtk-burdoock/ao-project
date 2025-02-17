extends Node

@export var create_scene_mobile: PackedScene  
@export var game_scene_mobile: PackedScene  
@export var create_scene_desktop: PackedScene  
@export var game_scene_desktop: PackedScene 

enum State{
	None,
	Login,
	Create	
}
 
@export var server_ip: LineEdit
@export var server_port: LineEdit
@export var music_active: Button
@export var label_error: Label
 
@onready var user_name:LineEdit = find_child("UserName")
@onready var user_password:LineEdit = find_child("UserPassword")
 
var _protocol:GameProtocol = null
var current_state:int = State.None

func _ready(): 
	Connection.connect("message_received", Callable(self, "_on_message_received"))
	Connection.connect("connected", Callable(self, "_on_client_connected"))
	Connection.connect("disconnected", Callable(self, "_on_client_disconnected"))
	Connection.connect("error_message",  Callable(self, "_on_error_message")) 

	_protocol.connect("logged", Callable(self, "_on_client_logged"))
	_protocol.connect("error_message", Callable(self, "_on_error_message")) 

func _on_BtnExit_pressed():
	get_tree().quit()
 
func _on_BtnConnect_pressed():
	if(current_state != State.None): return
	label_error.text = 'Conectando...'
	
	await get_tree().create_timer(0.5).timeout
	Configuration.server_ip = server_ip.text
	Configuration.server_port = server_port.text.to_int()
	Configuration.music = music_active.button_pressed
	await get_tree().create_timer(0.1).timeout
	current_state = State.Login
	var err = Connection.connect_to_server()	
	if err != 0:
		current_state = State.None
		await get_tree().create_timer(0.1).timeout
		_on_error_message('Error: '  + str(err) + '. IP Incorrecto')
	

func _on_BtnCreate_pressed():
	if(current_state != State.None): return
	await get_tree().create_timer(0.1).timeout
	label_error.text = 'Conectando...'
	
	Configuration.server_ip = server_ip.text
	Configuration.server_port = server_port.text.to_int()
	Configuration.music = music_active.button_pressed
	
	current_state = State.Create
	var err = Connection.connect_to_server()
	if err != 0:
		current_state = State.None
		await get_tree().create_timer(0.1).timeout
		_on_error_message('Error: '  + str(err) + '. IP Incorrecto')
	
func _on_message_received(data):
	_protocol.handle_incoming_data(data)
	
func _on_client_connected():	
	if(current_state == State.Create):
		var scene = create_scene_mobile.instantiate()		
		scene.initialize(_protocol)
		get_parent().switch_scene(scene)
	else:
		_protocol.write_login_existing_char(user_name.text, user_password.text)
		_protocol.flush_data()	
	print('Conexion establecida. Estado: ' , current_state)

func _on_client_disconnected():
	current_state = State.None
 
func _on_client_logged():	
	var scene: Node
	if Configuration.interface_mode == 0:		
		scene = game_scene_mobile.instantiate()
	else:
		scene = game_scene_desktop.instantiate()
	scene.initialize(_protocol)
	get_parent().switch_scene(scene)

func _on_error_message(message):
	label_error.text = str(message)

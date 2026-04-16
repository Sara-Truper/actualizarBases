package com.example.actualizarBases;
//método que se expone API o WS
//método que se expone API o WS

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.example.actualizarBases.Modelo.codigos;

@RestController
@RequestMapping("/act")
public class Controller {
	@Autowired
	private service mService;
	
	@PostMapping("/actualizar")
    public ResponseEntity<Map<String, Object>> actualizar() {
		Map<String, Object> respuesta = new HashMap<>();
        try {
            mService.actualizarProveedores();
            mService.actualizarDirectos();
            mService.actualizarPrecios();
            //mService.actualizarPool();
            respuesta.put("message", "Tablas (lista_proveedores, directos, precios)  actualizadas con éxito");
            return ResponseEntity.ok(respuesta);
        } catch (Exception e) {
            e.printStackTrace();
            respuesta.put("message", "Error al actualizar tablas: " + e.getMessage());
            return ResponseEntity.internalServerError().body(respuesta);
        }
    }
	
	@PostMapping("/pool")
	public ResponseEntity<Map<String, Object>> pool(){
		Map<String, Object> respuesta = new HashMap<>();
        try {
            mService.actualizarPool(); 
            respuesta.put("message", "Tabla Pool actualizada con éxito");
            return ResponseEntity.ok(respuesta);
        } catch (Exception e) {
            e.printStackTrace();
            respuesta.put("message", "Error al actualizar tabla: " + e.getMessage());
            return ResponseEntity.internalServerError().body(respuesta);
        }
	}
	
	@PostMapping("/actualizar2")
	public ResponseEntity<Map<String, Object>> contactos(){
		Map<String, Object> respuesta = new HashMap<>();
        try {
            mService.actualizarContactos();
            mService.actualizarWksh();
            mService.actualizarCodigos();
            respuesta.put("message", "Tablas (contactos, mksh, codigos) actualizadas con éxito");
            return ResponseEntity.ok(respuesta);
        } catch (Exception e) {
            e.printStackTrace();
            respuesta.put("message", "Error al actualizar tablas: " + e.getMessage());
            return ResponseEntity.internalServerError().body(respuesta);
        }
	}
}
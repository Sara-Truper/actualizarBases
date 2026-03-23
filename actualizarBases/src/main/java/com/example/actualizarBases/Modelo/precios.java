package com.example.actualizarBases.Modelo;

import java.math.BigDecimal;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.GeneratedValue;
import jakarta.persistence.GenerationType;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

@Entity
@Table(name="precios")
@AllArgsConstructor
@NoArgsConstructor
public class precios {
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	private Long Id;
	@Column(columnDefinition = "MEDIUMTEXT")
	private String materialproveedor;
	@Column(columnDefinition = "MEDIUMTEXT")
	private String material;
	@Column(columnDefinition = "MEDIUMTEXT")
	private String proveedor;
	@Column(name = "Precio", precision = 15, scale = 4)
	private BigDecimal Precio;
	private String moneda;
	public Long getId() {
		return Id;
	}
	public void setId(Long id) {
		Id = id;
	}
	public String getMaterialproveedor() {
		return materialproveedor;
	}
	public void setMaterialproveedor(String materialproveedor) {
		this.materialproveedor = materialproveedor;
	}
	public String getMaterial() {
		return material;
	}
	public void setMaterial(String material) {
		this.material = material;
	}
	public String getProveedor() {
		return proveedor;
	}
	public void setProveedor(String proveedor) {
		this.proveedor = proveedor;
	}
	public BigDecimal getPrecio() {
		return Precio;
	}
	public void setPrecio(BigDecimal precio) {
		Precio = precio;
	}
	public String getMoneda() {
		return moneda;
	}
	public void setMoneda(String moneda) {
		this.moneda = moneda;
	}
}

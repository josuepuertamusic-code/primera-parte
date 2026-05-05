import { Injectable } from '@angular/core';
import { Articulo } from '../models/articulo';

@Injectable({
  providedIn: 'root',
})
export class Articulos {
  private articulos: Articulo[] = [
    {
      id: 1,
      titulo: 'Correccion de errores del backend',
      resumen: 'Se ajustaron validaciones y respuestas HTTP.',
      autor: 'Josue',
      categoria: 'Backend',
      fecha: '2026-05-01',
      estado: 'publicado',
    },
    {
      id: 2,
      titulo: 'Pendiente: autenticacion JWT',
      resumen: 'Estructura inicial para iniciar sesion y proteger rutas.',
      autor: 'Equipo clase',
      categoria: 'Seguridad',
      fecha: '2026-05-03',
      estado: 'borrador',
    },
  ];

  getArticulos(): Articulo[] {
    return this.articulos;
  }

  addArticulo(nuevo: Omit<Articulo, 'id' | 'fecha'>): void {
    const articulo: Articulo = {
      ...nuevo,
      id: this.articulos.length + 1,
      fecha: new Date().toISOString().split('T')[0],
    };
    this.articulos = [articulo, ...this.articulos];
  }

  getTotalPublicados(): number {
    return this.articulos.filter((item) => item.estado === 'publicado').length;
  }
}

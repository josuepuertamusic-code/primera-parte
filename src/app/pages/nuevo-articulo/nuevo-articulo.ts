import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { RouterLink } from '@angular/router';
import { Articulos as ArticulosService } from '../../core/services/articulos';

@Component({
  selector: 'app-nuevo-articulo',
  imports: [FormsModule, RouterLink],
  templateUrl: './nuevo-articulo.html',
  styleUrl: './nuevo-articulo.css',
})
export class NuevoArticulo {
  model = {
    titulo: '',
    resumen: '',
    autor: '',
    categoria: 'General',
    estado: 'borrador' as const,
  };
  enviado = false;

  constructor(private articulosService: ArticulosService) {}

  guardar(): void {
    if (!this.model.titulo || !this.model.resumen || !this.model.autor) {
      return;
    }

    this.articulosService.addArticulo(this.model);
    this.enviado = true;
    this.model = {
      titulo: '',
      resumen: '',
      autor: '',
      categoria: 'General',
      estado: 'borrador',
    };
  }
}

import { Component } from '@angular/core';
import { Articulos as ArticulosService } from '../../core/services/articulos';

@Component({
  selector: 'app-dashboard',
  imports: [],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.css',
})
export class Dashboard {
  total = 0;
  publicados = 0;
  pendientes = 0;

  constructor(private articulosService: ArticulosService) {
    const lista = this.articulosService.getArticulos();
    this.total = lista.length;
    this.publicados = this.articulosService.getTotalPublicados();
    this.pendientes = this.total - this.publicados;
  }
}

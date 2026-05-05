import { Component } from '@angular/core';
import { DatePipe } from '@angular/common';
import { Articulos as ArticulosService } from '../../core/services/articulos';
import { Articulo } from '../../core/models/articulo';

@Component({
  selector: 'app-articulos',
  imports: [DatePipe],
  templateUrl: './articulos.html',
  styleUrl: './articulos.css',
})
export class Articulos {
  articulos: Articulo[] = [];

  constructor(private articulosService: ArticulosService) {
    this.articulos = this.articulosService.getArticulos();
  }
}

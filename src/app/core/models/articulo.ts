export interface Articulo {
  id: number;
  titulo: string;
  resumen: string;
  autor: string;
  categoria: string;
  fecha: string;
  estado: 'borrador' | 'publicado';
}

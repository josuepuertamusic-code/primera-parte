import { Routes } from '@angular/router';
import { Articulos } from './pages/articulos/articulos';
import { Dashboard } from './pages/dashboard/dashboard';
import { Inicio } from './pages/inicio/inicio';
import { NuevoArticulo } from './pages/nuevo-articulo/nuevo-articulo';

export const routes: Routes = [
  { path: '', component: Inicio, title: 'Inicio' },
  { path: 'dashboard', component: Dashboard, title: 'Dashboard' },
  { path: 'articulos', component: Articulos, title: 'Articulos' },
  { path: 'nuevo', component: NuevoArticulo, title: 'Nuevo articulo' },
  { path: '**', redirectTo: '' }
];

/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
*/

import React from 'react';

export interface Product {
  id: number; // agora é número
  name: string;
  tagline?: string; // pode ser opcional se não usar
  description: string;
  longDescription?: string;
  price: number;
  category:
    | "Carregadores de Celular"
    | "Comunicação Visual"
    | "Estruturas para Mídia OOH"
    | "Marcenaria"
    | "Projetos Especiais de Arquitetura"
    | "Totens Digitais e Interativos"; // novas categorias
  imageUrl: string;
  gallery?: string[];
  features?: string[];
}

export interface JournalArticle {
  id: number;
  title: string;
  date: string;
  excerpt: string;
  image: string;
  content: React.ReactNode;
}

export interface ChatMessage {
  role: 'user' | 'model';
  text: string;
  timestamp: number;
}

export enum LoadingState {
  IDLE = 'IDLE',
  LOADING = 'LOADING',
  ERROR = 'ERROR',
  SUCCESS = 'SUCCESS'
}

export type ViewState =
  | { type: 'home' }
  | { type: 'product'; product: Product }
  | { type: 'journal'; article: JournalArticle }
  | { type: 'checkout' };
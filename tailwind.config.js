/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  theme: {
    extend: {
      // Paleta tirada do próprio logo da Kalenborn: grafite #222220 e
      // amarelo #FFEC01. O índigo que estava aqui não era da marca.
      colors: {
        grafite: {
          50:'#F4F4F2', 100:'#E7E7E4', 200:'#CFCFCA', 300:'#A8A8A1',
          400:'#7A7A73', 500:'#55554F', 600:'#3B3B37', 700:'#2C2C29',
          800:'#222220', 900:'#191917', 950:'#0F0F0E',
        },
        // Acento. Não serve de fundo de texto nem de cor de texto em
        // branco: contraste ruim. Marca estado ativo e número-chave.
        marca: {
          DEFAULT:'#FFEC01', 600:'#D9C800', 700:'#A89B00',
        },
        papel: '#FAFAF7',
      },
      fontFamily: {
        // Plex foi desenhada para uma empresa de engenharia. A mono é
        // para identificadores — RAI, RAC, datas — que são códigos e se
        // leem melhor alinhados.
        sans: ['"IBM Plex Sans"', 'system-ui', 'sans-serif'],
        mono: ['"IBM Plex Mono"', 'ui-monospace', 'monospace'],
      },
      letterSpacing: { tightish: '-0.011em' },
    },
  },
  plugins: [],
}

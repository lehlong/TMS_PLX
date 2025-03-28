module.exports = {
  mode: 'jit',
  content: ['./src/**/*.{html,ts,scss}'],
  theme: {
    extend: {
      backgroundSize: {
        'size-100': '100% 100%',
      },
      colors: {
        'custom-gray': '#f0f2f5',
        validate: '#ff4d4f',
      },
      flexGrow: {
        2: '2',
        3: '3',
        4: '4',
      },
    },
  },
  plugins: [],
};

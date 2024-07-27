module.exports = {
    apps: [
      {
        name: 'incentive-app',
        script: 'app.py',
        interpreter: 'python',
        watch: true,
        env: {
          FLASK_ENV: 'development',
        },
        env_production: {
          FLASK_ENV: 'production',
        },
      },
    ],
  };
  
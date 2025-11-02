declare namespace NodeJS {
  interface ProcessEnv {
    MS_EMAIL: string;
    MS_PASSWORD: string;
    MS_TOTP_SECRET: string;
  }
}


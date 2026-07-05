# Ready Stock Parts B2B Portal Upgrade

## Files
- `app.py`: Professional B2B customer portal and admin dashboard.

## How to deploy on GitHub + Render
1. Rename `app_b2b_professional.py` to `app.py`.
2. Upload it to your GitHub repository and replace the old `app.py`.
3. Click **Commit changes**.
4. Go to Render.
5. Click **Manual Deploy → Deploy latest commit**.

## Admin password
Current admin password:

`Manutd@033`

For better security on Render, add an environment variable:

- Key: `ADMIN_PASSWORD`
- Value: your private password

The app will use the environment variable if it exists.

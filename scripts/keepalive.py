"""
keepalive.py — Mantém o app Streamlit acordado via navegador headless.

O Streamlit é uma SPA: um simples HTTP GET retorna 200 mas não acorda o app.
É preciso executar JavaScript e abrir o WebSocket, o que só um browser real faz.
"""

import asyncio
from playwright.async_api import async_playwright

URL = "https://receita-apsirtus.streamlit.app"


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        print(f"Visitando {URL} ...")
        await page.goto(URL, wait_until="domcontentloaded", timeout=120_000)
        await page.wait_for_timeout(5_000)

        # Se o app estiver dormindo, aparece o botão de acordar
        wake_btn = page.get_by_role("button", name="Yes, get this app back up!")
        if await wake_btn.count() > 0:
            print("  → App dormindo. Clicando em Wake up...")
            await wake_btn.click()
            await page.wait_for_timeout(60_000)
            print("  → App acordado!")
        else:
            print("  → App já estava acordado.")

        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())

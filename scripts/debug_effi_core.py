"""Debug effi-core connection."""
import asyncio
import traceback
from effi_work_client import get_client_identifiers_from_effi_work


async def main():
    try:
        print("Calling get_client_identifiers_from_effi_work...")
        result = await get_client_identifiers_from_effi_work("One2Call Limited")
        print(f"Result: {result}")
    except Exception as e:
        print(f"Exception: {e}")
        print(f"Type: {type(e)}")
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(main())

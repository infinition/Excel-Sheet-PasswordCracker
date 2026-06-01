# Excel Sheet Password Cracker

A VBA macro that removes the protection from a password-protected Excel
worksheet when you no longer have the password. It works by cycling through
character combinations until it finds one that Excel accepts as equivalent to
the original password, then unprotects the sheet with it.

This works because Excel's legacy sheet protection (pre-2010 `.xls` format and
the compatibility mode in newer files) uses a weak algorithm that maps many
different passwords to the same internal hash. The macro does not recover the
original password; it finds a string that produces the same hash and unprotects
the sheet with that.

For `.xlsx` files using the modern encryption standard this approach does not
apply.

---

## Usage

1. Open the Excel file with the protected worksheet.
2. Press `Alt+F11` to open the VBA editor, or go to Developer > View Code.
3. In the Project panel on the left, double-click the sheet that is protected.
4. Paste the contents of `Cracker.vba` into the code window that opens.
5. Press `F5` or click Run.

When a usable password is found, a dialog box appears:

```
One usable password is XXXYYYZZZXXXY
```

The sheet is now unprotected. The password shown is not necessarily the
original one; it is a functional equivalent that bypasses the check.

![Screenshot](https://user-images.githubusercontent.com/37984399/114324338-c1a96280-9b29-11eb-8f54-5355b7186c5d.png)

---

## How it works

The `PasswordBreaker` sub iterates over a fixed character range using twelve
nested loops, building 12-character strings from ASCII 65-66 (A and B) with a
final character sweeping the printable range (32-126). On each iteration it
calls `ActiveSheet.Unprotect` with the constructed string. If
`ActiveSheet.ProtectContents` returns `False`, the sheet accepted the password
and the macro exits.

The approach relies on Excel's legacy hash collisions. On a modern machine it
typically completes in seconds to a few minutes depending on where in the
iteration space the collision falls.

---

## Star History

<a href="https://www.star-history.com/?repos=infinition%2FExcel-Sheet-PasswordCracker&type=date&legend=top-left">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/chart?repos=infinition/Excel-Sheet-PasswordCracker&type=date&theme=dark&legend=top-left" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/chart?repos=infinition/Excel-Sheet-PasswordCracker&type=date&legend=top-left" />
   <img alt="Star History Chart" src="https://api.star-history.com/chart?repos=infinition/Excel-Sheet-PasswordCracker&type=date&legend=top-left" />
 </picture>
</a>

---

## Legal notice

Use this on files you own or have explicit authorisation to access. The author
takes no responsibility for misuse.

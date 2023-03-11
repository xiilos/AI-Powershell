

# Script #


function New-RandomPassword {
  param(
      [int]$Length = 12
  )

  $lowercase = "abcdefghijklmnopqrstuvwxyz"
  $uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  $numbers = "0123456789"
  $special = "!@#$%^&*()_-+=[]{};:,.<>?|"

  $allchars = $lowercase + $uppercase + $numbers + $special

  $password = ""
  for ($i = 0; $i -lt $Length; $i++) {
      $random = Get-Random -Minimum 0 -Maximum $allchars.Length
      $password += $allchars[$random]
  }
  return $password
}

# Generate a random password with default length of 12 characters
New-RandomPassword

# Generate a random password with a custom length of 16 characters
New-RandomPassword -Length 16



# End Scripting
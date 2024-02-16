# File Verification

This is a simple example how to organize verification of some ooxml_file.

## Steps

1. Install ruby
2. Install gem:

   ```bash
   gem install ooxml_parser
   ```

3. Run the script:

   ```ruby
    ruby file_verification.rb -f example.docx
   ```

In result it will output verification result,
also exit code will be 0 if file is valid, 1 if file is invalid.

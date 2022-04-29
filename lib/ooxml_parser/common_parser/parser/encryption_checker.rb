# frozen_string_literal: true

module OoxmlParser
  # Check if file is encrypted
  class EncryptionChecker
    def initialize(file_path, ignore_system: false)
      @file_path = file_path
      @ignore_system = ignore_system
    end

    # @return [Boolean] is file encrypted
    def encrypted?
      return false unless compatible_system?

      # Support of Encrypted status in `file` util was introduced in file v5.20
      # but LTS version of ubuntu before 16.04 uses older `file` and it return `Composite Document`
      # https://github.com/file/file/blob/0eb7c1b83341cc954620b45d2e2d65ee7df1a4e7/ChangeLog#L623
      if mime_type.include?('encrypted') ||
         mime_type.include?('Composite Document File V2 Document, No summary info') ||
         mime_type.include?('application/CDFV2-corrupt')
        warn("File #{@file_path} is encrypted. Can't parse it")
        return true
      end
      false
    end

    private

    # Check if system is compatible with `file` utility
    # @return [Boolean] Result of check
    def compatible_system?
      return true if !Gem.win_platform? && !@ignore_system

      warn 'Checking file for encryption is not supported on Windows'
      false
    end

    # @return [String] mime type of file
    def mime_type
      @mime_type ||= `file -b --mime "#{@file_path}"`
    end
  end
end

# frozen_string_literal: true

# Monkey patch around
# https://github.com/oracle/truffleruby/commit/cc76155bf509587e1b2954e9de77c558c7c857f4
# Remove it as soon as resolved
if defined?(Truffle)
  class File < IO
    SHARE_DELETE   = 0
  end
end

require_relative 'ooxml_parser/common_parser'
require_relative 'ooxml_parser/configuration'
require_relative 'ooxml_parser/helpers/string_helper'
require_relative 'ooxml_parser/docx_parser'
require_relative 'ooxml_parser/xlsx_parser'
require_relative 'ooxml_parser/pptx_parser'

# Basic class for all parser stuff
module OoxmlParser
  class << self
    attr_writer :configuration
  end

  def self.configuration
    @configuration ||= Configuration.new
  end

  # Reset settings to default
  def self.reset
    @configuration = Configuration.new
  end

  def self.configure
    yield(configuration)
  end
end

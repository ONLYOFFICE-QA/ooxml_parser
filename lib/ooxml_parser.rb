# frozen_string_literal: true

require_relative 'truffleruby_patch'
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

# frozen_string_literal: true

require 'securerandom'
require 'nokogiri'
require 'zip'
require 'ooxml_decrypt'
require_relative 'ooxml_document_object/nokogiri_parsing_exception'
require_relative 'ooxml_document_object/ooxml_document_object_helper'
require_relative 'ooxml_document_object/ooxml_object_attribute_helper'

module OoxmlParser
  # Basic class for any OOXML Document Object
  class OOXMLDocumentObject
    include OoxmlDocumentObjectHelper
    include OoxmlObjectAttributeHelper
    # @return [OOXMLDocumentObject] object which hold current object
    attr_accessor :parent

    def initialize(parent: nil)
      @parent = parent
    end

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      return false if self.class != other.class

      instance_variables.each do |current_attribute|
        next if current_attribute == :@parent
        next if instance_variable_get(current_attribute).is_a?(Nokogiri::XML::Element)

        return false unless instance_variable_get(current_attribute) == other.instance_variable_get(current_attribute)
      end
      true
    end

    # @return [True, false] if structure contain any user data
    def with_data?
      true
    end

    # @return [Nokogiri::XML::Document] result of parsing xml via nokogiri
    def parse_xml(xml_path)
      xml = Nokogiri::XML(File.open(xml_path), &:strict)
      unless xml.errors.empty?
        raise NokogiriParsingException,
              parse_error_message(xml_path, xml.errors)
      end
      xml
    rescue  Nokogiri::XML::SyntaxError => e
      raise NokogiriParsingException,
            parse_error_message(xml_path, e)
    end

    private

    # @param [String] xml_path path to xml
    # @param [String] errors errors
    # @return [String] error string
    def parse_error_message(xml_path, errors)
      "Nokogiri found errors in file: #{xml_path}. Errors: #{errors}"
    end
  end
end

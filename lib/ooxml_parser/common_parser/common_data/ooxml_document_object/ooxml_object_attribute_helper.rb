# frozen_string_literal: true

module OoxmlParser
  # Module for stroing some attribute helpers for objects
  module OoxmlObjectAttributeHelper
    # @param node [Nokogiri::XML:Element] node to parse
    # @param attribute_name [String] name of attribute
    # @return [True, False] is option enabled
    def option_enabled?(node, attribute_name = 'val')
      return true if node.attributes.empty?
      return true if node.to_s == '1'
      return false if node.to_s == '0'
      return false if node.attribute(attribute_name).nil?

      status = node.attribute(attribute_name).value
      !%w[false off 0].include?(status)
    end

    # @param node [Nokogiri::XML:Element] node to parse
    # @param attribute_name [String] name of attribute
    # @return [True, False] is attribute enabled
    def attribute_enabled?(node, attribute_name = 'val')
      return true if node.to_s == '1'
      return false if node.to_s == '0'
      return false if node.attribute(attribute_name).nil?

      status = node.attribute(attribute_name).value
      %w[true on 1].include?(status)
    end
  end
end

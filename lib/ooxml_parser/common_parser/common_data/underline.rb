# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `u` tags
  class Underline < OOXMLDocumentObject
    attr_accessor :style
    # @return [Color] color of underline
    attr_accessor :color
    # @return [Symbol] value of underline
    attr_reader :value

    def initialize(style = :none, color = nil, parent: nil)
      @style = style == 'single' ? :single : style
      @color = color
      super(parent: parent)
    end

    # Compare this object to other
    # @param other [Object] any other object
    # @return [True, False] result of comparision
    def ==(other)
      case other
      when Underline
        @style.to_sym == other.style.to_sym && @color == other.color
      when Symbol
        @style.to_sym == other
      else
        false
      end
    end

    # @return [String] result of convert of object to string
    def to_s
      if @color.nil?
        @style.to_s
      else
        "#{@style} #{@color}"
      end
    end

    # Parse Underline object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Underline] result of parsing
    def parse(node)
      parse_attributes(node) if node.is_a?(Nokogiri::XML::Element)

      case node
      when 'sng'
        @style = :single
      when 'none'
        @style = :none
      end
      self
    end

    private

    # Parse attributes
    # @param node [Nokogiri::XML:Element] node to parse
    def parse_attributes(node)
      node.attributes.each do |key, value|
        case key
        when 'color'
          @color = Color.new(parent: self).parse_hex_string(value.value)
        when 'val'
          @value = value.value.to_sym
        end
      end
    end
  end
end

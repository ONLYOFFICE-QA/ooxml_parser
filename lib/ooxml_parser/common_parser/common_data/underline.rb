# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `u` tags
  class Underline < OOXMLDocumentObject
    attr_accessor :style, :color

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
      case node
      when 'sng'
        @style = :single
      when 'none'
        @style = :none
      end
      self
    end
  end
end

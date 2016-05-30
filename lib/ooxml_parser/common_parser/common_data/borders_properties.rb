# Border Properties Data
module OoxmlParser
  class BordersProperties
    attr_accessor :color, :space, :sz, :val, :space, :shadow, :frame, :side

    def initialize(color = :auto, sz = 0, val = :none, space = 0)
      @color = color
      @sz = sz
      @val = val
      @space = space
    end

    def ==(other)
      instance_variables.each do |current_attributes|
        unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
          return false
        end
      end
      true
    end

    def nil?
      @sz == 0 && val == :none
    end

    def to_s
      return '' if nil?
      "borders color: #{@color}, size: #{@sz}, space: #{@space}, value: #{@val}"
    end

    def copy
      BordersProperties.new(@color, @sz, @val, @space)
    end

    def visible?
      return false if nil?
      !(val == 'none')
    end

    # Parse BordersProperties
    # @param [Nokogiri::XML:Element] node with BordersProperties
    # @return [BordersProperties] value of BordersProperties
    def self.parse(node)
      return nil if node.attribute('val').value == 'nil'
      border_properties = BordersProperties.new
      border_properties.val = node.attribute('val').value.to_sym
      border_properties.sz = node.attribute('sz').value.to_f / 8.0 if node.attribute('sz')
      border_properties.space = (node.attribute('space').value.to_f / 28.34).round(3) unless node.attribute('space').nil?
      if node.attribute('color')
        border_properties.color = node.attribute('color').value
        unless node.attribute('shadow').nil?
          border_properties.shadow = node.attribute('shadow').value
        end
        if border_properties.color != 'auto'
          border_properties.color = Color.from_int16(border_properties.color)
        end
      end
      border_properties
    end
  end
end

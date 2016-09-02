module OoxmlParser
  # Size of some object
  class OoxmlSize
    # @return [Float] value of size
    attr_accessor :value
    # @return [Symbol] units of measurement
    attr_accessor :unit

    def initialize(value = 0, unit = :dxa)
      @unit = unit
      @value = value
    end

    # Parse OoxmlSize
    # @param [Nokogiri::XML:Node] node with OoxmlSize
    # @return [OoxmlSize] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'type'
          @unit = value.value.to_sym
        when 'w'
          @value = value.value.to_f
        end
      end
      self
    end

    def ==(other)
      (to_base_unit.value - other.to_base_unit.value).abs < 10**(OoxmlParser.configuration.accuracy + 2)
    end

    def zero?
      @value.zero?
    end

    def to_s
      "#{@value} #{@unit}"
    end

    # Convert all values to one same base unit
    # @return [OoxmlSize] base unit
    def to_base_unit
      case @unit
      when :centimeter
        return OoxmlSize.new(@value * 360_000)
      when :point
        return OoxmlSize.new(@value * 12_700)
      when :half_point
        return OoxmlSize.new(@value * (12_700 / 2))
      when :one_eighth_point
        return OoxmlSize.new(@value * (12_700 / 8))
      when :dxa, :twip
        return OoxmlSize.new(@value * 635, :emu)
      when :inch
        return OoxmlSize.new(@value * 914_400, :emu)
      else
        return self
      end
    end

    # @param output_unit [Symbol] output unit of convertion
    # @return [OoxmlSize] converted unit
    def to_unit(output_unit)
      base_unit = to_base_unit
      case output_unit
      when :centimeter
        return OoxmlSize.new((base_unit.value / 360_000).round(OoxmlParser.configuration.accuracy), output_unit)
      when :point
        return OoxmlSize.new((base_unit.value / 12_700).round(OoxmlParser.configuration.accuracy), output_unit)
      when :half_point
        return OoxmlSize.new((base_unit.value / (12_700 * 2)).round(OoxmlParser.configuration.accuracy), output_unit)
      when :one_eighth_point
        return OoxmlSize.new((base_unit.value / (12_700 * 8)).round(OoxmlParser.configuration.accuracy), output_unit)
      when :dxa, :twip
        return OoxmlSize.new((base_unit.value / 635).round(OoxmlParser.configuration.accuracy), output_unit)
      when :inch
        return OoxmlSize.new((base_unit.value / 914_400).round(OoxmlParser.configuration.accuracy), output_unit)
      when :percent
        return OoxmlSize.new((base_unit.value / 50).round(OoxmlParser.configuration.accuracy), output_unit)
      else
        return base_unit
      end
    end
  end
end

module OoxmlParser
  # Size of some object
  class OoxmlSize
    # @return [Float] value of size
    attr_accessor :value
    # @return [Symbol] units of measurement
    attr_accessor :unit

    def initialize(value, unit = :dxa)
      @unit = unit
      @value = value
    end

    def ==(other)
      (to_emu.value - other.to_emu.value).abs < 10**(OoxmlParser.configuration.accuracy + 2)
    end

    def zero?
      @value.zero?
    end

    def to_s
      "#{@value} #{@unit}"
    end

    def to_emu
      case @unit
      when :centimeter
        return OoxmlSize.new(@value * 360_000, @unit)
      when :point
        return OoxmlSize.new(@value * 12_700, @unit)
      when :half_point
        return OoxmlSize.new(@value * (12_700 * 2), @unit)
      when :dxa, :twip
        return OoxmlSize.new(@value * 635, @unit)
      when :inch
        return OoxmlSize.new(@value * 914_400, @unit)
      when :emu
        return self
      end
    end

    def to_unit(output_unit)
      emu = to_emu
      case output_unit
      when :centimeter
        return OoxmlSize.new((emu.value / 360_000).round(OoxmlParser.configuration.accuracy), output_unit)
      when :point
        return OoxmlSize.new((emu.value / 12_700).round(OoxmlParser.configuration.accuracy), output_unit)
      when :half_point
        return OoxmlSize.new((emu.value / (12_700 * 2)).round(OoxmlParser.configuration.accuracy), output_unit)
      when :dxa, :twip
        return OoxmlSize.new((emu.value / 635).round(OoxmlParser.configuration.accuracy), output_unit)
      when :inch
        return OoxmlSize.new((emu.value / 914_400).round(OoxmlParser.configuration.accuracy), output_unit)
      when :emu
        return emu
      end
    end
  end
end

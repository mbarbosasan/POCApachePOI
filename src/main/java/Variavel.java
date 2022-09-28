public class Variavel {
    private String elementoEP;
    private String grandeza;
    private String tipo;
    private String fonte;
    private String servidor;
    private String tag;
    private String unidadeFonte;

    public Variavel(String elementoEP, String grandeza, String tipo, String fonte, String servidor, String tag, String unidadeFonte) {
        this.elementoEP = elementoEP;
        this.grandeza = grandeza;
        this.tipo = tipo;
        this.fonte = fonte;
        this.servidor = servidor;
        this.tag = tag;
        this.unidadeFonte = unidadeFonte;
    }

    @Override
    public String toString() {
        return "Variavel{" +
                "elementoEP='" + elementoEP + '\'' +
                ", grandeza='" + grandeza + '\'' +
                ", tipo='" + tipo + '\'' +
                ", fonte='" + fonte + '\'' +
                ", servidor='" + servidor + '\'' +
                ", tag='" + tag + '\'' +
                ", unidadeFonte='" + unidadeFonte + '\'' +
                '}';
    }
}

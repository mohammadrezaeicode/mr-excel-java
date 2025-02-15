package io.github.mohammadrezaeicode.excel.model;

public class Comment {
    private String comment;
    private String styleId;
    private String author;

    public Comment(String comment, String styleId, String author) {
        this.comment = comment;
        this.styleId = styleId;
        this.author = author;
    }

    public String getComment() {
        return comment;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }

    public String getStyleId() {
        if(styleId==null){
            return "";
        }
        return styleId;
    }

    public void setStyleId(String styleId) {
        this.styleId = styleId;
    }

    public String getAuthor() {
        if(author==null){
            return "";
        }
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }
}

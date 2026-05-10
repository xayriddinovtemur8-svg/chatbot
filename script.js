const canvas = document.getElementById("canvas");
const ctx = canvas.getContext("2d");

canvas.width = window.innerWidth;
canvas.height = window.innerHeight;

window.addEventListener("resize", () => {
    canvas.width = window.innerWidth;
    canvas.height = window.innerHeight;
});

let mouse = {
    x: canvas.width / 2,
    y: canvas.height / 2
};

window.addEventListener("mousemove", (e) => {
    mouse.x = e.x;
    mouse.y = e.y;
});

class Segment {
    constructor(x, y, size) {
        this.x = x;
        this.y = y;
        this.size = size;
    }

    draw() {

        // Body
        ctx.beginPath();
        ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
        ctx.fillStyle = "rainbow";
        ctx.fill();

        // Legs
        ctx.beginPath();
        ctx.moveTo(this.x - this.size, this.y);
        ctx.lineTo(this.x - this.size - 10, this.y - 8);

        ctx.moveTo(this.x + this.size, this.y);
        ctx.lineTo(this.x + this.size + 10, this.y + 8);

        ctx.strokeStyle = "rgba(255,255,255,0.5)";
        ctx.lineWidth = 1;
        ctx.stroke();
    }
}

const segments = [];
const total = 45;

for(let i = 0; i < total; i++) {
    segments.push(
        new Segment(
            mouse.x,
            mouse.y,
            8 - i * 0.12
        )
    );
}

function animate() {

    ctx.clearRect(0, 0, canvas.width, canvas.height);

    let x = mouse.x;
    let y = mouse.y;

    segments.forEach((seg, index) => {

        seg.x += (x - seg.x) * 0.35;
        seg.y += (y - seg.y) * 0.35;

        seg.draw();

        // Body line
        ctx.beginPath();
        ctx.moveTo(seg.x, seg.y);

        if(index !== 0){
            ctx.lineTo(
                segments[index - 1].x,
                segments[index - 1].y
            );
        }

        ctx.strokeStyle = "rgba(255,255,255,0.35)";
        ctx.lineWidth = 2;
        ctx.stroke();

        x = seg.x;
        y = seg.y;
    });

    requestAnimationFrame(animate);
}

animate();
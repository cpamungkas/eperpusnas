from fastapi import FastAPI, Depends, HTTPException, status
from jose import JWTError, jwt
from passlib.context import CryptContext
from datetime import datetime, timedelta

# Secret key buat encode/decode token
SECRET_KEY = "secret-key-yang-super-rahasia" # Ganti dengan secret yang aman
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 30  # Token akan expired dalam 30 menit

app = FastAPI()

# Setup hashing password
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# Simulasi database user
fake_users_db = {
    "user": {
        "username": "user",
        "hashed_password": pwd_context.hash("password"), # "password" ini contoh aja
    }
}

# Fungsi buat verifikasi password
def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

# Fungsi buat autentikasi user
def authenticate_user(username: str, password: str):
    user = fake_users_db.get(username)
    if not user or not verify_password(password, user["hashed_password"]):
        return False
    return user

# Fungsi buat create JWT token
def create_access_token(data: dict, expires_delta: timedelta | None = None):
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.utcnow() + expires_delta
    else:
        expire = datetime.utcnow() + timedelta(minutes=15)
    to_encode.update({"exp": expire})
    encoded_jwt = jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)
    return encoded_jwt

# Endpoint buat login dan dapetin token
@app.post("/login")
def login(username: str, password: str):
    user = authenticate_user(username, password)
    if not user:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Incorrect username or password")
    
    access_token = create_access_token(data={"sub": user["username"]}, expires_delta=timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES))
    return {"access_token": access_token, "token_type": "bearer"}

# Fungsi buat verify JWT token
from fastapi.security import OAuth2PasswordBearer

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="login")

def get_current_user(token: str = Depends(oauth2_scheme)):
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username = payload.get("sub")
        if username is None:
            raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid token")
        return username
    except JWTError:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid token")

# Endpoint protected, cuma bisa diakses dengan JWT
@app.get("/api/preview")
def preview_file(file_path: str, current_user: str = Depends(get_current_user)):
    decoded_path = urllib.parse.unquote(file_path)
    if os.path.exists(decoded_path):
        return FileResponse(decoded_path)
    else:
        raise HTTPException(status_code=404, detail="File not found")
